import pandas as pd
import oracledb
from datetime import datetime
import os
import sys
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Adicionar raiz do projecto ao path para importar utilitários partilhados
sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection

# ── Receber data via argumento ────────────────────────────────
if len(sys.argv) < 3:
    print("❌ Uso: python pagamentos_cartao.py YYYYMMDD PASTA_DESTINO")
    sys.exit(1)

data_ref = sys.argv[1]

# ── Validação básica do formato ───────────────────────────────
try:
    datetime.strptime(data_ref, "%Y%m%d")
except ValueError:
    print(f"❌ Data inválida: '{data_ref}'. Use o formato YYYYMMDD.")
    sys.exit(1)

print(f"📅 Data de referência: {data_ref}")


excel_path = "C:/Users/sum001781/Desktop/projeto/templates/Mapa de Reporte de Pagamentos via Cartão.xlsx"
#excel_path = "T:/Direcção de Informação de Gestão/16. Data Desk/4. Data Desk/BdM/file py/templates/Mapa de Reporte de Pagamentos via Cartão.xlsx"
# Query — sem alterações em relação ao script original
query = f"""
SELECT /*+ PARALLEL */ NUIB,"Número do cartão","Tipo de Cartão","Número de conta",NIB,NUIT,"Nome do Titular","Tipo de terminal", "Detalhe de terminal","Rede de pagamento","País da terminal",
    "Moeda da Transação","Montante em moeda estrangeira","Montante em USD","Montante em moeda Nacional","Data da operação","Data da liquidação"
FROM (
WITH EDST AS (
    SELECT LOCALIZ_TERM, COD_PAIS, b.moeda, 
    CASE WHEN DEB_CRD_VALOR = 'C' THEN VALOR * -1 ELSE VALOR END AS VALOR, 
    TO_DATE(SUBSTR(A.DTH_TRANS,1,8),'YYYYMMDD') AS DIA_TRANSACAO, 
    DATA_LIQUIDACAO, NUM_CARTAO, CONTRATO, 
        CASE
          WHEN
            (A.TIPO_TERM='C' AND A.COD_TRANS IN ('01', '31')) OR
            (A.TIPO_TERM='D' AND A.COD_TRANS IN ('10', '11', '15', '16', '34', '39'))
          THEN TO_NUMBER(SUBSTR(A.DADOS_ADICIONAIS, 2, 3))
        END COD_MOEDA_INFO_COMP,
        CASE
            WHEN A.TIPO_TERM IN ('A','C') THEN 'ATM'
            WHEN A.TIPO_TERM IN ('B','D') THEN 'POS'
            WHEN A.TIPO_TERM = 'M' THEN 'INTERNET'
        END AS DESC_TIPO_TERMINAL,
        CASE
          WHEN
            (TIPO_TERM='C' AND COD_TRANS IN ('01', '31')) OR
            (TIPO_TERM='D' AND COD_TRANS IN ('10', '11', '15', '16', '34', '39'))
          THEN TO_NUMBER(SUBSTR(DADOS_ADICIONAIS, 14, 11) )/100
        END MONTANTE_ESTRANGEIRO, 
        A.NRLOG
    FROM HUB.MOVS_EDST1 A
    LEFT JOIN DWMB.LKP_MOEDA B ON B.ISO = A.COD_MOEDA_OPER
    WHERE DATA_LIQUIDACAO = TO_DATE({data_ref}, 'YYYYMMDD')
      AND COD_PAIS <> 508
),
NUIB AS (
    SELECT ICNTIT, ICVCOP
    FROM (
        SELECT ICNTIT, ICVCOP, rank () over (partition by ICNTIT order by ICDTCR desc, dtref desc, ICHOCR desc) rk
        FROM HUB.IFCOPC
        WHERE ICCCOP = 'NUIB' AND ICSITU = 'V'
    )
    WHERE rk = 1
),
NUIT AS (
    SELECT COD_ENT, NUIT
    FROM (
        SELECT  
            NCDCNUMT AS COD_ENT,
            NCDCNDOC AS NUIT,
            RANK() OVER (
                PARTITION BY NCDCNUMT
                ORDER BY NCDCDATACR DESC, NCDCHORACR DESC, NCDCTPDOC
            ) AS RNK
        FROM HUB.NCDOCENT
        WHERE DTREF = {data_ref}
          AND TRIM(NCDCTPDOC) IN ('ENI','EN1','EN2','EN3','EN4','EN5','EN6','EN7','EN8','EN9')
          AND NCDCSITU = 'V'
    )
    WHERE RNK = 1
)
SELECT DISTINCT /*+ PARALLEL */
    COALESCE(T3.NUIB, TT4.ICVCOP) AS NUIB,
    E.PAN "Número do cartão",
    CASE
        WHEN TRIM(E.COMPONENTE) LIKE '%PP%' THEN 'PRÉ-PAGO'
        WHEN TRIM(E.COMPONENTE) NOT LIKE '%PP%' AND TRIM(E.CLASSE) = 'CARD' THEN 'DÉBITO'  
        WHEN TRIM(E.CLASSE) = 'CARC' THEN 'CRÉDITO'
    END AS "Tipo de Cartão",
    T2.COD_CONTRATO "Número de conta",
    D2.NIB,
    COALESCE(T3.NUIT, T4.NUIT) AS NUIT,
    T3.NOME AS "Nome do Titular",
    A.DESC_TIPO_TERMINAL "Tipo de terminal", 
	A.LOCALIZ_TERM "Detalhe de terminal",
	'VISA' "Rede de pagamento",
	C.DSC "País da terminal",
	B.MOEDA "Moeda da Transação",
    A.MONTANTE_ESTRANGEIRO "Montante em moeda estrangeira",
    ABS(CASE 
            WHEN  B.MOEDA = 'USD' THEN A.MONTANTE_ESTRANGEIRO 
            ELSE ROUND (( (A.VALOR / 100)  * (-1) / FX.TX_COMPRA), 2)
        END
    ) AS "Montante em USD",
    A.VALOR / 100 "Montante em moeda Nacional",
	A.DIA_TRANSACAO "Data da operação",
	A.DATA_LIQUIDACAO "Data da liquidação",
    A.NRLOG
FROM EDST A
LEFT JOIN DWMB.LKP_PAIS C ON C.ISO = A.COD_PAIS
LEFT JOIN DWMB.LKP_MOEDA B ON B.ISO = A.COD_MOEDA_INFO_COMP
LEFT JOIN DWMB.T_FAC_CAMBIO FX
           ON FX.DT_COT =  A.DATA_LIQUIDACAO
          AND FX.COD_MOEDA_BASE = 'MZN'
          AND FX.COD_MOEDA = 'USD'
          AND FX.COD_TIPO_COT = 'R'
LEFT JOIN DWMB.DIM_CARTAO E ON SUBSTR(E.PAN,1,15) = SUBSTR(A.NUM_CARTAO,1,15)
LEFT JOIN DWMB.T_DIM_CONTRATO D2 ON D2.COD_CONTRATO = E.CONTRATO_DO_1 AND TO_DATE ({data_ref}, 'YYYYMMDD') BETWEEN D2.DT_INIC_REG AND D2.DT_FIM_REG
LEFT JOIN DWMB.LKP_CONTRATO T2 ON T2.COD_CONTRATO = D2.COD_CONTRATO
LEFT JOIN DWMB.LKP_ENTIDADE T3 ON T2.COD_ENT = T3.COD_ENT
LEFT JOIN NUIB TT4 ON T3.COD_ENT = TT4.ICNTIT
LEFT JOIN NUIT T4 ON T4.COD_ENT = T3.COD_ENT
WHERE T2.COD_CONTRATO > 0
) 

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

if 'NUIB' in df.columns:
    df['NUIB'] = df['NUIB'].astype(str).str.zfill(15)
if 'Número do cartão' in df.columns:
    df['Número do cartão'] = df['Número do cartão'].astype(str).str.zfill(16)
if 'NIB' in df.columns:
    df['NIB'] = df['NIB'].astype(str).str.zfill(21)
if 'NUIT' in df.columns:
    df['NUIT'] = df['NUIT'].astype(str).replace('nan', '')

# ── Sem dados: não gera ficheiro (assim o botão de download não aparece) ──
if df.empty:
    print("⚠️ A consulta não retornou resultados. Nenhum ficheiro foi gerado.")
    sys.exit(0)

# ── Abrir template e preencher (openpyxl — sem Excel a correr) ───────────
wb = load_workbook(excel_path)
ws = wb.worksheets[0]

# Preencher campos do cabeçalho
data_min = pd.to_datetime(df['Data da liquidação'].min()).strftime('%Y-%m-%d')
ws["D6"] = data_min
ws["D7"] = data_min
ws["D6"].alignment = Alignment(horizontal="left")
ws["D7"].alignment = Alignment(horizontal="left")

# Ordem das colunas conforme o template (C → S)
colunas_template = [
    'NUIB',                            # C
    'Número do cartão',                # D
    'Tipo de Cartão',                  # E
    'Número de conta',                 # F
    'NIB',                             # G
    'NUIT',                            # H
    'Nome do Titular',                 # I
    'Tipo de terminal',                # J
    'Detalhe de terminal',             # K
    'Rede de pagamento',               # L
    'País da terminal',                # M
    'Moeda da Transação',              # N
    'Montante em moeda estrangeira',   # O
    'Montante em USD',                 # P
    'Montante em moeda Nacional',      # Q
    'Data da operação',                # R
    'Data da liquidação'               # S
]

colunas_faltantes = [col for col in colunas_template if col not in df.columns]
if colunas_faltantes:
    print(f"⚠️ Colunas faltantes: {colunas_faltantes}")

df_template = df[colunas_template].copy()
df_template['Data da operação']    = pd.to_datetime(df_template['Data da operação']).dt.strftime('%Y-%m-%d')
df_template['Data da liquidação']  = pd.to_datetime(df_template['Data da liquidação']).dt.strftime('%Y-%m-%d')

# Colunas que devem ficar como texto (para não virarem notação científica) — índices 0-based
COLUNAS_TEXTO = {0, 1, 3, 4, 5}  # NUIB, Número do cartão, Número de conta, NIB, NUIT

print("\n📝 A carregar dados no Excel...")
dados_para_excel = df_template.values

for i, linha in enumerate(dados_para_excel):
    linha_excel = i + 12  # Começar na linha 12
    for j, valor in enumerate(linha):
        coluna_excel = j + 3  # Começar na coluna C (3)
        cell = ws.cell(row=linha_excel, column=coluna_excel)
        if j in COLUNAS_TEXTO:
            # Ao contrário do xlwings, aqui não é preciso o truque do apóstrofe —
            # basta escrever o valor como string e marcar o formato como texto ("@").
            cell.value = str(valor)
            cell.number_format = "@"
        else:
            cell.value = valor

print("✅ Dados inseridos!")

# Formatar colunas de montante como números (O, P, Q = colunas 15, 16, 17)
for row in range(12, 12 + len(df_template)):
    for col in [15, 16, 17]:
        cell = ws.cell(row=row, column=col)
        if cell.value is not None:
            cell.number_format = '#,##0.00'

# Garantir que a coluna NIB (G = 7) inteira fica como texto
for row in range(12, 12 + len(df_template)):
    ws.cell(row=row, column=7).number_format = "@"

# Extrair nome base do ficheiro de template
nome_base = os.path.splitext(os.path.basename(excel_path))[0]
nome_ficheiro = f"{nome_base}_{data_ref}.xlsx"

pasta_destino = sys.argv[2]
caminho_final = os.path.join(pasta_destino, nome_ficheiro)

try:
    wb.save(caminho_final)
except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)

print(f"✅ Ficheiro guardado em: {caminho_final}")