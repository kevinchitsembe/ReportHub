import sys
import os
import pandas as pd
import oracledb
from datetime import datetime
import calendar
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection
from openpyxl.utils import get_column_letter
import unicodedata
from decimal import Decimal, ROUND_HALF_UP, ROUND_FLOOR
# ── Argumentos ────────────────────────────────
if len(sys.argv) < 3:
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM PASTA_DESTINO")
    sys.exit(1)

data_ref = sys.argv[1]
pasta_destino = sys.argv[2]
try:
    datetime.strptime(data_ref, "%Y%m%d")
except ValueError:
    print(f"❌ Data inválida: '{data_ref}'. Use o formato YYYYMMDD.")
    sys.exit(1)

print(f"📅 Data de referência: {data_ref}")

#excel_path = "T:/Direcção de Informação de Gestão/16. Data Desk/4. Data Desk/BdM/file py/templates/Mapa de Depósitos por Genero Sector e Distrito.xlsx"
excel_path = "C:/Users/sum001781/Desktop/projeto/templates/Mapa de Depósitos por Genero Sector e Distrito.xlsx"
# ── Queries ────────────────────────────────────
# CORRIGIDO (face à v2): nas queries originais (4 queries), o filtro
# "COD_SEG<>'CI' AND COD_SIT NOT IN('8','9') AND COD_CNT_CONTAB<>'991'"
# só se aplicava ao VOLUME (Query 1/2) — na query de VALOR (Query 3/4),
# esse filtro estava preso dentro de um OR que não o aplicava aos
# depósitos. A v2 juntou tudo num único WHERE e acabou a aplicar esse
# filtro também ao Valor, o que reduzia indevidamente os valores.
# Agora o WHERE externo só identifica "é um depósito"
# (COD_CLASSE_COMP_RNA + COD_CNT_CONTAB LIKE '40%'), e o filtro extra
# entra só dentro do COUNT(DISTINCT CASE WHEN ...), via FILTRO_VOLUME —
# nunca dentro do SUM(...), reproduzindo fielmente a assimetria original.
#
# ⚠️ Confirmar com o DBA/modelo de dados:
#   - PROVINCIA / BDM_NOME assumidos como vindos de C (AGENCIAS_ASOF_202101)
#   - COD_SEXO / COD_EMP_PART / DSC_SEGMENTOCOMERCIAL / COD_CAE assumidos
#     como vindos de B (T_DIM_ENT)

SQL_GENERO = """
SELECT SK_TEMP_SLD_CNT, PROVINCIA, BDM_NOME,
       COUNT(DISTINCT CASE WHEN SEXO = 'MULHERES' AND FILTRO_VOLUME = 1 THEN COD_CONTRATO END) MULHERES_QTD,
       COUNT(DISTINCT CASE WHEN SEXO = 'HOMENS'   AND FILTRO_VOLUME = 1 THEN COD_CONTRATO END) HOMENS_QTD,
       COUNT(DISTINCT CASE WHEN SEXO = 'OUTROS'   AND FILTRO_VOLUME = 1 THEN COD_CONTRATO END) OUTROS_QTD,
       NVL(SUM(CASE WHEN SEXO = 'MULHERES' THEN ABS(VAL_SLD_CONV) END), 0) MULHERES_VAL,
       NVL(SUM(CASE WHEN SEXO = 'HOMENS'   THEN ABS(VAL_SLD_CONV) END), 0) HOMENS_VAL,
       NVL(SUM(CASE WHEN SEXO = 'OUTROS'   THEN ABS(VAL_SLD_CONV) END), 0) OUTROS_VAL
FROM (
    SELECT E.SK_TEMP_SLD_CNT,
           CASE WHEN A.COD_UNID_NEG >= 1000 THEN 'Maputo' ELSE C.PROVINCIA END PROVINCIA,
           CASE WHEN A.COD_UNID_NEG >= 1000 THEN 'Cidade de Maputo' ELSE C.BDM_NOME END BDM_NOME,
           CASE WHEN B.COD_SEXO = 'F' THEN 'MULHERES'
                WHEN B.COD_SEXO = 'M' THEN 'HOMENS'
                ELSE 'OUTROS' END SEXO,
           A.COD_CONTRATO,
           E.VAL_SLD_CONV,
           -- Este filtro só se aplicava ao VOLUME nas queries originais
           -- (na query de Valor, os depósitos entravam sem esta condição,
           -- por causa do OR que isolava esse filtro do lado dos "créditos").
           CASE WHEN E.COD_SEG <> 'CI'
                 AND E.COD_SIT NOT IN ('8','9')
                 AND E.COD_CNT_CONTAB <> '991'
                THEN 1 ELSE 0 END FILTRO_VOLUME
    FROM DWMB.T_FAC_SLD_CNT E
    LEFT JOIN DWMB.T_DIM_CONTRATO A ON E.SK_CONTRATO = A.SK_CONTRATO
    LEFT JOIN DWMB.T_DIM_ENT B ON E.SK_ENT = B.SK_ENT
    LEFT JOIN X900065.AGENCIAS_ASOF_202101 C ON A.COD_UNID_NEG = C.CC
    WHERE E.SK_TEMP_SLD_CNT = :data_ref
      AND E.COD_CLASSE_COMP_RNA IN ('DO', 'DP')
      AND E.COD_CNT_CONTAB LIKE '40%'
)
GROUP BY SK_TEMP_SLD_CNT, PROVINCIA, BDM_NOME
"""

SQL_SETOR = """
SELECT SK_TEMP_SLD_CNT, PROVINCIA, BDM_NOME,
       COUNT(DISTINCT CASE WHEN TIPO_CLIENTE = 'SINGULARES'  AND FILTRO_VOLUME = 1 THEN COD_CONTRATO END) SINGULARES_QTD,
       COUNT(DISTINCT CASE WHEN TIPO_CLIENTE = 'MPMES'       AND FILTRO_VOLUME = 1 THEN COD_CONTRATO END) MPMES_QTD,
       COUNT(DISTINCT CASE WHEN TIPO_CLIENTE = 'AGRICULTURA' AND FILTRO_VOLUME = 1 THEN COD_CONTRATO END) AGRICULTURA_QTD,
       COUNT(DISTINCT CASE WHEN TIPO_CLIENTE = 'OUTROS'      AND FILTRO_VOLUME = 1 THEN COD_CONTRATO END) OUTROS_QTD,
       NVL(SUM(CASE WHEN TIPO_CLIENTE = 'SINGULARES'  THEN ABS(VAL_SLD_CONV) END), 0) SINGULARES_VAL,
       NVL(SUM(CASE WHEN TIPO_CLIENTE = 'MPMES'       THEN ABS(VAL_SLD_CONV) END), 0) MPMES_VAL,
       NVL(SUM(CASE WHEN TIPO_CLIENTE = 'AGRICULTURA' THEN ABS(VAL_SLD_CONV) END), 0) AGRICULTURA_VAL,
       NVL(SUM(CASE WHEN TIPO_CLIENTE = 'OUTROS'      THEN ABS(VAL_SLD_CONV) END), 0) OUTROS_VAL
FROM (
    SELECT E.SK_TEMP_SLD_CNT,
           CASE WHEN A.COD_UNID_NEG >= 1000 THEN 'Maputo' ELSE C.PROVINCIA END PROVINCIA,
           CASE WHEN A.COD_UNID_NEG >= 1000 THEN 'Cidade de Maputo' ELSE C.BDM_NOME END BDM_NOME,
           CASE WHEN SUBSTR(B.COD_CAE, 1, 2) IN ('01') THEN 'AGRICULTURA'
                WHEN B.COD_EMP_PART = 'E' AND F.DSC_SEGMENTOCOMERCIAL = 'Retalho' THEN 'MPMES'
                WHEN B.COD_EMP_PART = 'P' AND A.COD_COMP_RNA = 'DL_MPESA' THEN 'OUTROS'
                WHEN B.COD_EMP_PART = 'P' THEN 'SINGULARES'
                ELSE 'OUTROS' END TIPO_CLIENTE,
           A.COD_CONTRATO,
           E.VAL_SLD_CONV,
           CASE WHEN E.COD_SEG <> 'CI'
                 AND E.COD_SIT NOT IN ('8','9')
                 AND E.COD_CNT_CONTAB <> '991'
                THEN 1 ELSE 0 END FILTRO_VOLUME
    FROM DWMB.T_FAC_SLD_CNT E
    LEFT JOIN DWMB.T_DIM_CONTRATO A ON E.SK_CONTRATO = A.SK_CONTRATO
    LEFT JOIN DWMB.T_DIM_ENT B ON E.SK_ENT = B.SK_ENT
    LEFT JOIN X900065.AGENCIAS_ASOF_202101 C ON A.COD_UNID_NEG = C.CC
    LEFT JOIN DWMB.DIM_ORGANIZACIONAL F ON A.COD_UNID_NEG=F.UNIDADENEGOCIO
    WHERE E.SK_TEMP_SLD_CNT = :data_ref
      AND E.COD_CLASSE_COMP_RNA IN ('DO', 'DP')
      AND E.COD_CNT_CONTAB LIKE '40%'
)
GROUP BY SK_TEMP_SLD_CNT, PROVINCIA, BDM_NOME
"""

# ── Executar as 2 queries ─────────────────────
print("⏳ A executar as consultas (versão otimizada: 2 queries em vez de 4)...")

try:
    with connection() as conn:
        cursor = conn.cursor()

        def fetch_df(sql, params):
            cursor.execute(sql, params)
            cols = [c[0] for c in cursor.description]
            return pd.DataFrame(cursor.fetchall(), columns=cols)

        params = {"data_ref": int(data_ref)}
        df_genero = fetch_df(SQL_GENERO, params)   # Volume + Valor - Género
        df_setor = fetch_df(SQL_SETOR, params)     # Volume + Valor - Sector

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

def normalizar(nome):
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return "OUTROS"  # saldos sem agência/contrato contam na linha OUTROS
    s = unicodedata.normalize("NFKD", str(nome)).encode("ascii", "ignore").decode()
    return " ".join(s.split()).upper()

def preparar_dict(df, colunas_valor):
    if df.empty:
        return {}

    df = df.copy()

    # Converte cada coluna de valor para Decimal, tratando NaN como Decimal("0")
    for coluna in colunas_valor:
        df[coluna] = df[coluna].map(
            lambda valor: Decimal("0") if pd.isna(valor) else Decimal(str(valor))
        )

    df["CHAVE"] = df["BDM_NOME"].map(normalizar)

    return (
        df.groupby("CHAVE", dropna=False)[colunas_valor]
        .sum()
        .to_dict("index")
    )

def arredondar_inteiro(valor):
    if valor is None or pd.isna(valor):
        return 0
    valor = valor if isinstance(valor, Decimal) else Decimal(str(valor))
    return int(valor.quantize(Decimal("1"), rounding=ROUND_HALF_UP))

def arredondar_preservando_soma(valores):
    """valores: {chave: Decimal}. Devolve {chave: int} em que a soma dos
    inteiros == round(soma dos Decimals) (método do maior resto)."""
    alvo = int(sum(valores.values(), Decimal("0")).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    base = {k: int(v.to_integral_value(rounding=ROUND_FLOOR)) for k, v in valores.items()}
    em_falta = alvo - sum(base.values())
    por_resto = sorted(valores, key=lambda k: valores[k] - base[k], reverse=True)
    for k in por_resto[:em_falta]:
        base[k] += 1
    return base

COLS_GENERO_QTD = ['MULHERES_QTD', 'HOMENS_QTD', 'OUTROS_QTD']
COLS_GENERO_VAL = ['MULHERES_VAL', 'HOMENS_VAL', 'OUTROS_VAL']
COLS_SETOR_QTD = ['SINGULARES_QTD', 'MPMES_QTD', 'AGRICULTURA_QTD', 'OUTROS_QTD']
COLS_SETOR_VAL = ['SINGULARES_VAL', 'MPMES_VAL', 'AGRICULTURA_VAL', 'OUTROS_VAL']

dados_genero = preparar_dict(df_genero, COLS_GENERO_QTD + COLS_GENERO_VAL)
dados_setor = preparar_dict(df_setor, COLS_SETOR_QTD + COLS_SETOR_VAL)

# ── Mapeamento do template ────────────────────
LINHAS_TEMPLATE = [
    (15, "Cidade de Maputo", "Cidade de Maputo"),
    (16, "Cidade da Matola", "Cidade da Matola"),
    (17, "Boane",            "Boane"),
    (18, "Magude",           "Magude"),
    (19, "Manhiça",          "Manhiça"),
    (20, "Marracuene",       "Marracuene"),
    (21, "Matutuíne",        "Matutuíne"),
    (22, "Namaacha",         "Namaacha"),
    (23, "Moamba",           "Moamba"),
    (25, "Cidade de Xai-Xai", "Cidade de Xai-Xai"),
    (26, "Bilene-Macia",      "Bilene-Macia"),
    (27, "Chibuto",           "Chibuto"),
    (28, "Chicualacuala",     "Chicualacuala"),
    (29, "Chigubo",           "Chigubo"),
    (30, "Chókwè",            "Chókwè"),
    (31, "Guijá",             "Guijá"),
    (32, "Mabalane",          "Mabalane"),
    (33, "Manjacaze",         "Manjacaze"),
    (34, "Massangena",        "Massangena"),
    (35, "Massingir",         "Massingir"),
    (36, "Chonguene",         "Chonguene"),
    (37, "Mapai",             "Mapai"),
    (38, "Limpopo",           "Limpopo"),
    (40, "Cidade de Inhambane", "Cidade de Inhambane"),
    (41, "Cidade de Maxixe",    "Cidade de Maxixe"),
    (42, "Funhaloro",           "Funhaloro"),
    (43, "Govuro",              "Govuro"),
    (44, "Homoine",             "Homoine"),
    (45, "Inharrime",           "Inharrime"),
    (46, "Inhassoro",           "Inhassoro"),
    (47, "Jangamo",             "Jangamo"),
    (48, "Mabote",              "Mabote"),
    (49, "Massinga",            "Massinga"),
    (50, "Morrumbene",          "Morrumbene"),
    (51, "Panda",               "Panda"),
    (52, "Vilankulo",           "Vilankulo"),
    (53, "Zavala (Quissico)",   "Zavala (Quissico)"),
    (55, "Cidade da Beira",  "Cidade da Beira"),
    (56, "Buzi",             "Buzi"),
    (57, "Caia",             "Caia"),
    (58, "Chemba",           "Chemba"),
    (59, "Cheringoma",       "Cheringoma"),
    (60, "Chibabava",        "Chibabava"),
    (61, "Dondo",            "Dondo"),
    (62, "Gorongoza",        "Gorongoza"),
    (63, "Machanga",         "Machanga"),
    (64, "Maringue",         "Maringue"),
    (65, "Marromeu",         "Marromeu"),
    (66, "Muanza",           "Muanza"),
    (67, "Nhamatanda",       "Nhamatanda"),
    (69, "Cidade de Chimoio", "Cidade de Chimoio"),
    (70, "Bárue",             "Bárue"),
    (71, "Gondola",           "Gondola"),
    (72, "Guro",              "Guro"),
    (73, "Machaze",           "Machaze"),
    (74, "Macossa",           "Macossa"),
    (75, "Manica",            "Manica"),
    (76, "Mossurize",         "Mossurize"),
    (77, "Sussundenga",       "Sussundenga"),
    (78, "Tambara",           "Tambara"),
    (79, "Macate",            "Macate"),
    (80, "Vanduzi",           "Vanduzi"),
    (82, "Cidade de Tete",   "Cidade de Tete"),
    (83, "Angónia",          "Angónia"),
    (84, "Cahora-Bassa",     "Cahora-Bassa"),
    (85, "Changara",         "Changara"),
    (86, "Chifunde",         "Chifunde"),
    (87, "Chiuta",           "Chiuta"),
    (88, "Macanga",          "Macanga"),
    (89, "Mágoe",            "Mágoe"),
    (90, "Maravia",          "Maravia"),
    (91, "Moatize",          "Moatize"),
    (92, "Mutarara",         "Mutarara"),
    (93, "Tsangano",         "Tsangano"),
    (94, "Zumbo",            "Zumbo"),
    (95, "Doa",              "Doa"),
    (96, "Marara",           "Marara"),
    (98,  "Cidade de Quelimane", "Cidade de Quelimane"),
    (99,  "Alto Molócuè",        "Alto Molócuè"),
    (100, "Chinde",              "Chinde"),
    (101, "Gurúè",               "Gurúè"),
    (102, "Ile",                 "Ile"),
    (103, "Inhassunge",          "Inhassunge"),
    (104, "Lugela",              "Lugela"),
    (105, "Gilé",                "Gilé"),
    (106, "Maganja da Costa",    "Maganja da Costa"),
    (107, "Milange",             "Milange"),
    (108, "Mocuba",              "Mocuba"),
    (109, "Mopeia",              "Mopeia"),
    (110, "Morrumbala",          "Morrumbala"),
    (111, "Namacurra",           "Namacurra"),
    (112, "Namarroi",            "Namarroi"),
    (113, "Nicoadala",           "Nicoadala"),
    (114, "Pebane",              "Pebane"),
    (115, "Molumbo",             "Molumbo"),
    (116, "Mocubela",            "Mocubela"),
    (117, "Mulevala",            "Mulevala"),
    (118, "Luabo",               "Luabo"),
    (119, "Derre",               "Derre"),
    (121, "Cidade de Nampula",       "Cidade de Nampula"),
    (122, "Cidade de Nacala-Porto",  "Cidade de Nacala-Porto"),
    (123, "Angoche",                 "Angoche"),
    (124, "Ilha de Moçambique",      "Ilha de Moçambique"),
    (125, "Lalaua",                  "Lalaua"),
    (126, "Malema",                  "Malema"),
    (127, "Meconta (Namialo)",       "Meconta (Namialo)"),
    (128, "Mecuburi",                "Mecuburi"),
    (129, "Memba",                   "Memba"),
    (130, "Mogincual",               "Mogincual"),
    (131, "Mogovolas",               "Mogovolas"),
    (132, "Moma",                    "Moma"),
    (133, "Monapo",                  "Monapo"),
    (134, "Mossuril",                "Mossuril"),
    (135, "Muecate",                 "Muecate"),
    (136, "Murrupula",               "Murrupula"),
    (137, "Nacala-Velha",            "Nacala-Velha"),
    (138, "Nacaroa",                 "Nacaroa"),
    (139, "Namapa-Erati",            "Namapa-Erati"),
    (140, "Rapale-Nampula",          "Rapale-Nampula"),
    (141, "Ribaué",                  "Ribaué"),
    (142, "Liupo",                   "Liupo"),
    (143, "Larde",                   "Larde"),
    (145, "Cidade de Pemba",        "Cidade de Pemba"),
    (146, "Ancuabe",                "Ancuabe"),
    (147, "Balama",                 "Balama"),
    (148, "Chiúre",                 "Chiúre"),
    (149, "Ibo",                    "Ibo"),
    (150, "Macomia",                "Macomia"),
    (151, "Mecufi",                 "Mecufi"),
    (152, "Meluco",                 "Meluco"),
    (153, "Mocimboa da Praia",      "Mocimboa da Praia"),
    (154, "Montepuez",              "Montepuez"),
    (155, "Mueda",                  "Mueda"),
    (156, "Muidumbe",               "Muidumbe"),
    (157, "Namuno",                 "Namuno"),
    (158, "Nangade",                "Nangade"),
    (159, "Palma",                  "Palma"),
    (160, "Pemba-Metuge",           "Pemba-Metuge"),
    (161, "Quissanga",              "Quissanga"),
    (163, "Cidade de Lichinga", "Cidade de Lichinga"),
    (164, "Cuamba",             "Cuamba"),
    (165, "Lago",               "Lago"),
    (166, "Manjune",            "Manjune"),
    (167, "Mandimba",           "Mandimba"),
    (168, "Marrupa",            "Marrupa"),
    (169, "Maúa",               "Maúa"),
    (170, "Mavago",             "Mavago"),
    (171, "Mecanhelas",         "Mecanhelas"),
    (172, "Mecula",             "Mecula"),
    (173, "Metarica",           "Metarica"),
    (174, "Muembe",             "Muembe"),
    (175, "N´gauma",            "N´gauma"),
    (176, "Nipepe",             "Nipepe"),
    (177, "Sanga",              "Sanga"),
    (178, "Chimbunhila",        "Chimbunhila"),
]

# Linhas de soma por província
LINHA_Maputo         = 14
LINHA_Gaza           = 24
LINHA_Inhambane      = 39
LINHA_Sofala         = 54
LINHA_Manica         = 68
LINHA_Tete           = 81
LINHA_Zambezia       = 97
LINHA_Nampula        = 120
LINHA_Cabo_Delgado   = 144
LINHA_Niassa         = 162
LINHA_Total          = 179

# ── Diagnóstico: o que a query trouxe mas não tem linha no template ──
chaves_template = {normalizar(c) for _, _, c in LINHAS_TEMPLATE}
for nome, dados, cols in [("GÉNERO", dados_genero, COLS_GENERO_VAL),
                          ("SECTOR", dados_setor,  COLS_SETOR_VAL)]:
    for chave, linha in dados.items():
        if chave not in chaves_template:
            print(f"⚠️ {nome}: '{chave}' sem linha no template -> "
                  + ", ".join(f"{c}={linha[c]:,.0f}" for c in cols))
    tot_bd = {c: sum(l[c] for l in dados.values()) for c in cols}
    print(f"ℹ️ Total {nome} na BD: {tot_bd}")

# ── Preencher o Excel ─────────────────────────
try:
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.worksheets[0]
    ws.data_validations = None
except Exception as e:
    print(f"❌ Erro ao carregar template Excel: {e}")
    sys.exit(1)

data_ref_dt  = datetime.strptime(data_ref, "%Y%m%d")
primeiro_dia = datetime(data_ref_dt.year, data_ref_dt.month, 1)
ultimo_dia   = datetime(
    data_ref_dt.year,
    data_ref_dt.month,
    calendar.monthrange(data_ref_dt.year, data_ref_dt.month)[1]
)

ws["B3"] = str(data_ref_dt.year)
ws["B3"].number_format = "@"
ws["B4"].number_format = "@"
ws["B4"] = primeiro_dia.strftime('%Y-%m-%d')
ws["B5"].number_format = "@"
ws["B5"] = ultimo_dia.strftime('%Y-%m-%d')

print("📝 A carregar dados no Excel...")

FONTE_NORMAL = Font(bold=False)

# Colunas (1-based no Excel) -> nome da coluna calculada nas queries
COL_VOL_GEN = {'MULHERES_QTD': 2, 'HOMENS_QTD': 3, 'OUTROS_QTD': 4}
COL_VOL_SEC = {'SINGULARES_QTD': 5, 'MPMES_QTD': 6, 'AGRICULTURA_QTD': 7, 'OUTROS_QTD': 8}
COL_VAL_GEN = {'MULHERES_VAL': 9, 'HOMENS_VAL': 10, 'OUTROS_VAL': 11}
COL_VAL_SEC = {'SINGULARES_VAL': 12, 'MPMES_VAL': 13, 'AGRICULTURA_VAL': 14, 'OUTROS_VAL': 15}

COLUNAS_VOLUME = set(COL_VOL_GEN.values()) | set(COL_VOL_SEC.values())   # 2..8
COLUNAS_VALOR  = set(COL_VAL_GEN.values()) | set(COL_VAL_SEC.values())   # 9..15


def get_valor(dados_dict, bdm_nome, coluna_nome):
    """Devolve o valor já agregado (lookup em dict, O(1)) para a agência
    e coluna."""
    linha = dados_dict.get(normalizar(bdm_nome))
    if linha is None:
        return 0
    return float(linha.get(coluna_nome, 0) or 0)


def set_cell(linha, col_excel, valor, number_format):
    """Escreve numa célula, convertendo 0 (ou None) em '-'."""
    c = ws.cell(row=linha, column=col_excel)
    if valor is None or float(valor) == 0:
        c.value = "-"
        c.alignment = Alignment(horizontal='center')
    else:
        c.value = arredondar_inteiro(valor)  # arredondado, tal como pedido
    c.font = FONTE_NORMAL
    c.number_format = number_format


def escrever_linha(linha, bdm_nome):
    """Escreve os 14 valores de uma agência numa linha do Excel."""
    for nome_col, col_excel in COL_VOL_GEN.items():
        set_cell(linha, col_excel, get_valor(dados_genero, bdm_nome, nome_col), '#,##0')
    for nome_col, col_excel in COL_VOL_SEC.items():
        set_cell(linha, col_excel, get_valor(dados_setor, bdm_nome, nome_col), '#,##0')
    for nome_col, col_excel in COL_VAL_GEN.items():
        set_cell(linha, col_excel, get_valor(dados_genero, bdm_nome, nome_col), '#,##0')
    for nome_col, col_excel in COL_VAL_SEC.items():
        set_cell(linha, col_excel, get_valor(dados_setor, bdm_nome, nome_col), '#,##0')


# Preencher linhas individuais
for linha, nome_template, chave_bd in LINHAS_TEMPLATE:
    escrever_linha(linha, chave_bd)


def somar_linhas(linhas_origem):
    """Soma os valores já escritos nas células (lê de volta do Excel, não do
    DataFrame) — trata '-' como 0."""
    totais = {c: 0 for c in COLUNAS_VOLUME | COLUNAS_VALOR}
    for l in linhas_origem:
        for c in totais:
            v = ws.cell(row=l, column=c).value
            if v is not None and v != "-":
                totais[c] += float(v)
    return totais


def escrever_totais(linha, totais):
    """Escreve uma linha de totais, usando o mesmo set_cell() (0 -> '-')."""
    for col, v in totais.items():
        set_cell(linha, col, v, '#,##0')


# ── Somatórios por província ──────────────────
GRUPOS_PROVINCIA = [
    (LINHA_Maputo, range(15, 23 + 1)),
    (LINHA_Gaza, range(25, 38 + 1)),
    (LINHA_Inhambane, range(40, 53 + 1)),
    (LINHA_Sofala, range(55, 67 + 1)),
    (LINHA_Manica, range(69, 80 + 1)),
    (LINHA_Tete, range(82, 96 + 1)),
    (LINHA_Zambezia, range(98, 119 + 1)),
    (LINHA_Nampula, range(121, 143 + 1)),
    (LINHA_Cabo_Delgado, range(145, 161 + 1)),
    (LINHA_Niassa, range(163, 178 + 1)),
]

for linha_total_prov, linhas_distritos in GRUPOS_PROVINCIA:
    totais = somar_linhas(list(linhas_distritos))
    escrever_totais(linha_total_prov, totais)

# ── Linha Total = soma das linhas de somatório provincial ──
LINHAS_PROVINCIAIS = [linha for linha, _ in GRUPOS_PROVINCIA]
totais_gerais = somar_linhas(LINHAS_PROVINCIAIS)

# Saldos sem agência/contrato: não têm distrito, entram só na linha Total
CHAVE_SEM_AG = normalizar(None)
extra_sem_agencia = {}
# Valores: somar os dados sem arredondamento e arredondar apenas no Total.
# A soma inclui também as chaves sem linha no template, como "OUTROS".
for dados, mapa_val in [
    (dados_genero, COL_VAL_GEN),
    (dados_setor,  COL_VAL_SEC),
]:
    brutos = {
        nome_col: sum((l.get(nome_col, Decimal("0")) for l in dados.values()), Decimal("0"))
        for nome_col in mapa_val
    }
    arred = arredondar_preservando_soma(brutos)
    for nome_col, col_excel in mapa_val.items():
        totais_gerais[col_excel] = arred[nome_col]

escrever_totais(LINHA_Total, totais_gerais)

# ── Validação cruzada (opcional, mas recomendada) ──
totais_distritos = somar_linhas([l for l, _, _ in LINHAS_TEMPLATE])
divergencias = {
    col: (totais_distritos[col], totais_gerais[col])
    for col in totais_distritos
    if abs(totais_distritos[col] + extra_sem_agencia.get(col, 0) - totais_gerais[col]) > 0.01
}
if divergencias:
    print("⚠️  Divergências entre soma de distritos e linha Total:")
    for col, (a, b) in divergencias.items():
        print(f"   Coluna {col}: distritos={a}  total={b}  diff={a-b}")
else:
    print("✅ Validação OK: soma dos distritos == linha Total.")

print("✅ Dados inseridos!")

# ── Ajustar largura das colunas B..O (2..15) ──
for col_idx in range(2, 16):   # B a O
    letra = get_column_letter(col_idx)
    max_len = 0
    for row_idx in range(1, ws.max_row + 1):
        v = ws.cell(row=row_idx, column=col_idx).value
        if v is None or v == "-":
            continue
        try:
            s = f"{float(v):,.0f}"
        except (ValueError, TypeError):
            s = str(v)
        max_len = max(max_len, len(s))
    ws.column_dimensions[letra].width = max(10, max_len + 2)

# ── Guardar ───────────────────────────────────
nome_base = os.path.splitext(os.path.basename(excel_path))[0]
nome_ficheiro = f"{nome_base}_{data_ref}.xlsx"
caminho_final = os.path.join(pasta_destino, nome_ficheiro)

try:
    wb.save(caminho_final)
except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)

print(f"✅ Ficheiro guardado em: {caminho_final}")