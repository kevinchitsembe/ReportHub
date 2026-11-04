import pandas as pd
import oracledb
from datetime import datetime
import os
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import calendar
# Adicionar raiz do projecto ao path para importar utilitários partilhados
sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection

# ── Receber data via argumento ────────────────────────────────
if len(sys.argv) < 3:
    print("❌ Uso: python pagamentos_cartao.py YYYYMMDD PASTA_DESTINO")
    sys.exit(1)

data_ref = sys.argv[1]
pasta_destino = sys.argv[2]
# ── Validação básica do formato ───────────────────────────────
try:
    datetime.strptime(data_ref, "%Y%m%d")
except ValueError:
    print(f"❌ Data inválida: '{data_ref}'. Use o formato YYYYMMDD.")
    sys.exit(1)

print(f"📅 Data de referência: {data_ref}")

#"T:/Direcção de Informação de Gestão/16. Data Desk/4. Data Desk/BdM/file py/templates/Mapa de Reporte de Pagamentos via Cartão.xlsx"
excel_path = "C:/Users/sum001781/Desktop/projeto/templates/Mapa 5a - Contas Bancárias.xlsx"

# Query — sem alterações em relação ao script original
query = f"""
WITH NUIB AS 

(
select ICNCLI, IFICCONTA, ICNTIT, ICVCOP, ICDTCR 
from (
SELECT ICNCLI, IFICCONTA, ICNTIT, ICVCOP, ICDTCR, rank () over (partition by ICNTIT order by ICDTCR desc, dtref desc, ICHOCR desc) rk
FROM HUB.IFCOPC
WHERE ICCCOP = 'NUIB' AND ICSITU = 'V'
)
where rk = 1
),

CONTAS_REPORT_ESTADO AS 
    (
        SELECT IFICCONTA, 1 FLAG_ESTADO
        FROM
        (
            SELECT rank() over (PARTITION by IFICCONTA ORDER BY ICDTCR desc, ICHOCR desc,ICDTDE desc, ICHODE desc) AS RANK,A.*
            FROM HUB.IFCOPC A 
            where A.ICCCOP='TSTAT' 
            ORDER BY  IFICCONTA, ICDTCR, ICHOCR,ICDTDE , ICHODE
        )
        WHERE RANK=1 AND ICSITU='V' 
            
    )

SELECT NIB,coalesce (T.NUIB, ICVCOP ) "NUIB do Titular", CASE WHEN nib is null then null else concat ('MZ59 ', T1.NIB) end IBAN, T4.DSC_UNIDADENEGOCIO "Agência / ID Ponto", 'N/A' "Agente / ID Ponto ",
 CASE   WHEN T.COD_EMP_PART = 'E' THEN '05'  
        WHEN (T1.COD_CLASSE_COMP = 'DO' AND T1.COD_COMP_RNA = 'DO_BAS23') THEN '02'
        WHEN (T1.COD_CLASSE_COMP = 'DO' AND ((T1.COD_COMP <> 'DO_BAS23' AND  T1.COD_COMP_RNA not IN ('DO_SALAR21','DO_SALPROT', 'DO_SALARI0','DS_SALARI0' ) OR UPPER(T1.COD_COMP_RNA) not LIKE '%SAL_R%')
        AND T.COD_EMP_PART <> 'E')) THEN '06'
        WHEN T1.COD_CLASSE_COMP = 'DP' THEN '03'
        WHEN (T1.COD_CLASSE_COMP in( 'DO', 'DS') AND (T1.COD_COMP_RNA  IN ('DO_SALAR21','DO_SALPROT', 'DO_SALARI0','DS_SALARI0' ) OR UPPER(T1.COD_COMP_RNA) LIKE '%SAL_R%')) THEN '04'
        WHEN T1.COD_CLASSE_COMP NOT IN ('DO', 'DP', 'DS') THEN '07'
      END "Tipo de conta", T1.COD_MOEDA Moeda, 
 CASE WHEN T.COD_TIPO_ENT IN ( 'P', 'A','N', 'M', 'W' ) AND TT1.COD_TIPO_TIT = 'I' THEN '01'
      WHEN T.COD_TIPO_ENT IN ( 'P', 'A','N', 'M', 'W' ) AND TT1.COD_TIPO_TIT = 'C' THEN '02'
      WHEN T.COD_TIPO_ENT IN ( 'P', 'A','N', 'M', 'W' ) AND TT1.COD_TIPO_TIT = 'S' THEN '03'
      WHEN T.COD_TIPO_ENT IN ( 'E', 'I','F' ) THEN '04'
      WHEN T.COD_TIPO_ENT IN ( 'E', 'I' ) AND TT2.FLAG_ESTADO IS NOT NULL THEN '05'
      WHEN T.COD_TIPO_ENT IN ( 'E', 'I','F' ) AND COD_CAE = '93021' THEN '06'
      WHEN T.COD_TIPO_ENT IN ( 'E', 'I','F' ) AND (COD_CAE = '64191'  or (nome like '%CR_DITO%' OR nome like '%BANCO%' OR nome like '%BANK%' OR nome like '%SEGURO%')) THEN '07'
      WHEN T.COD_TIPO_ENT IN ('F') THEN '08' 
      ELSE 'N/A'
      END  Titularidade,
 CASE WHEN T.COD_TIPO_ENT = 'M' THEN '01' WHEN T.COD_TIPO_ENT IN ('E', 'F') AND COD_EMP_PART = 'E' THEN '03' ELSE '02' END "Tipo de Titular",
 case when COD_COND_MOV = '10' then '01'
      when COD_COND_MOV = '50' then '02'
      when COD_COND_MOV in ('20', '30', '40', '60') then '03'
      when COD_COND_MOV = '999' then '05' else 'N/A'
 end "Regime de movimentação", t.dsc_sector "Sector de Actividade do Titular", 
'N/A' "Dimensão da Empresa titular", 'N/A' "Volume de Negócios da Empresa titular", 'N/A' "Número de trabalhadores da empresa titular", 'N/A' "Tipo de classificação da necessidade especial/deficiência",
'N/A' "Canal preferencial de acesso à conta", 
 CASE WHEN T1.COD_SIT in ('N', '1', '2', '3', '4', '5', '6', '7', '8', '9') THEN '01'
      WHEN T1.COD_SIT = 'I' THEN '02'
      WHEN T1.COD_SIT = 'E' THEN '04'
      ELSE 'N/A' 
      END "Estado da Conta",
'N/A' "Motivo de alteração do Estado", T1.DT_SIT "Data da última actualização do estado",
t1.DT_ABERT "Data de Abertura", CASE WHEN T1.COD_SIT = 'E' THEN T1.DT_SIT ELSE NULL END "Data de Encerramento",t3.DATA_ULTIMO_MOVIMENTO_CLIENTE "Data da Última Operação", 
t.DT_ACT "Data da Última Actualização dos dados", '01' "Canal de Abertura",
'N/A'  "Frequência de utilização da conta",  CASE WHEN TT1.COD_SEG = 'BS' THEN 'S' ELSE 'G' END "Regime de Abertura",
CASE WHEN FLAG_PEP = 'S' THEN 'S' ELSE 'N' END "Pessoa Politicamente Exposta (PEP)"

FROM DWMB.lkp_ENTidade T
JOIN DWMB.lkp_CONTRATO T1 ON T.cod_ENT = T1.cod_ENT --AND  &DTRF BETWEEN T1.DT_INIc_REG AND T1.DT_FIM_REG
JOIN DWMB.lkp_CoNTa TT1 ON T.cod_ENT = TT1.cod_ENT-- AND  &DTRF BETWEEN TT1.DT_INIc_REG AND TT1.DT_FIM_REG
LEFT JOIN NUIB T2 ON T.COD_ENT = ICNTIT
LEFT JOIN CONTAS_REPORT_ESTADO TT2 ON T1.COD_CONTRATO = TT2.IFICCONTA
LEFT JOIN DWMB.T_FAC_SLD_CNT T3 ON T1.COD_CONTRATO = T3.COD_CONTRATO 
LEFT JOIN DWMB.DIM_ORGANIZACIONAL T4 ON T1.COD_UNID_NEG = T4.UNIDADENEGOCIO
WHERE SK_temp_sld_cnt = {data_ref}
"""

# ── Executar query via pool ───────────────────
print("⏳ A executar consulta...")

try:
    with connection() as conn:
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [col[0] for col in cursor.description]
        data    = cursor.fetchall()
        df      = pd.DataFrame(data, columns=columns)

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

if 'NUIB do Titular' in df.columns:
    df['NUIB do Titular'] = df['NUIB do Titular'].astype(str).str.zfill(15)
if 'NIB' in df.columns:
    df['NIB'] = df['NIB'].astype(str).str.zfill(21)

# ── Sem dados: não gera ficheiro (assim o botão de download não aparece) ──
if df.empty:
    print("⚠️ A consulta não retornou resultados. Nenhum ficheiro foi gerado.")
    sys.exit(0)

try:
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.worksheets[0]
    #print("✅ Template carregado com sucesso.")
    
    # REMOVER TODAS AS VALIDAÇÕES DE DADOS (DROPDOWNS)
    ws.data_validations = None
    #print("✅ Validações de dados (dropdowns) removidas.")
    
except Exception as e:
    print(f"❌ Erro ao carregar template Excel: {e}")
    sys.exit(1)

hoje = datetime.today()

# Primeiro dia do mês anterior
if hoje.month == 1:
    primeiro_dia = datetime(hoje.year - 1, 12, 1)
else:
    primeiro_dia = datetime(hoje.year, hoje.month - 1, 1)

# Último dia do mês anterior
ultimo_dia = datetime(
    primeiro_dia.year,
    primeiro_dia.month,
    calendar.monthrange(primeiro_dia.year, primeiro_dia.month)[1]
)

# Preencher cabeçalho
ws["C4"] = str(hoje.year)
ws["C4"].number_format = "@"

ws["C5"].number_format = "@"
ws["C5"] = primeiro_dia.strftime('%Y-%m-%d')

ws["C6"].number_format = "@"
ws["C6"] = ultimo_dia.strftime('%Y-%m-%d')

# Ordem das colunas conforme o template (C → S)
colunas_template = [
    'NIB',                            # C
    'NUIB do Titular',                # D
    'IBAN',                  # E
    'Agência / ID Ponto',                 # F
    'Agente / ID Ponto ',                             # G
    'Tipo de conta',                            # H
    'MOEDA',                 # I
    'TITULARIDADE',                # J
    'Tipo de Titular',             # K
    'Regime de movimentação',               # L
    'Sector de Actividade do Titular',                # M
    'Dimensão da Empresa titular',              # N
    'Volume de Negócios da Empresa titular',   # O
    'Número de trabalhadores da empresa titular',                 # P
    'Tipo de classificação da necessidade especial/deficiência',      # Q
    'Canal preferencial de acesso à conta',                # R
    'Estado da Conta',               # S
    'Motivo de alteração do Estado',
    'Data da última actualização do estado',
    'Data de Abertura',
    'Data de Encerramento',
    'Data da Última Operação',
    'Data da Última Actualização dos dados',
    'Canal de Abertura',
    'Frequência de utilização da conta',
    'Regime de Abertura',
    'Pessoa Politicamente Exposta (PEP)',
]

colunas_faltantes = [col for col in colunas_template if col not in df.columns]
if colunas_faltantes:
    print(f"⚠️ Colunas faltantes: {colunas_faltantes}")

df_template = df[colunas_template].copy()
df_template['Data da última actualização do estado']    = pd.to_datetime(df_template['Data da última actualização do estado']).dt.strftime('%d-%m-%Y')
df_template['Data de Abertura']  = pd.to_datetime(df_template['Data de Abertura']).dt.strftime('%d-%m-%Y')
df_template['Data de Encerramento']  = pd.to_datetime(df_template['Data de Encerramento']).dt.strftime('%d-%m-%Y')
df_template['Data da Última Operação']  = pd.to_datetime(df_template['Data da Última Operação']).dt.strftime('%d-%m-%Y')
df_template['Data da Última Actualização dos dados']  = pd.to_datetime(df_template['Data da Última Actualização dos dados']).dt.strftime('%d-%m-%Y')
# Colunas que devem ficar como texto (para não virarem notação científica) — índices 0-based
COLUNAS_TEXTO = {0, 1}  # NUIB, Número do cartão, NIB, NUIT

FONTE_NORMAL = Font(bold=False)  # reset explícito — não herda negrito do template

print("\n📝 A carregar dados no Excel...")
dados_para_excel = df_template.values

for i, linha in enumerate(dados_para_excel):
    linha_excel = i + 11  # Começar na linha 12
    for j, valor in enumerate(linha):
        coluna_excel = j + 1  # Começar na coluna A
        cell = ws.cell(row=linha_excel, column=coluna_excel)
        cell.value = valor
        cell.font = FONTE_NORMAL  # limpa qualquer negrito residual do template

        if j in COLUNAS_TEXTO:
            cell.value = str(valor)
            cell.number_format = "@"
            

print("✅ Dados inseridos!")

# Garantir que a coluna NIB (G = 7) inteira fica como texto
#for row in range(12, 12 + len(df_template)):
   # ws.cell(row=row, column=7).number_format = "@"

# Extrair nome base do ficheiro de template
nome_base = os.path.splitext(os.path.basename(excel_path))[0]
nome_ficheiro = f"{nome_base}_{data_ref}.xlsx"


caminho_final = os.path.join(pasta_destino, nome_ficheiro)

try:
    wb.save(caminho_final)
except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)

#print(f"✅ Ficheiro guardado em: {caminho_final}")