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
if len(sys.argv) < 4:
    print("❌ Uso: python pagamentos_cartao.py YYYYMMDD PASTA_DESTINO")
    sys.exit(1)

data_ref = sys.argv[1]
ficheiro = sys.argv[2]
pasta_destino = sys.argv[3]

# ── Validação básica do formato ───────────────────────────────
try:
    datetime.strptime(data_ref, "%Y%m%d")
except ValueError:
    print(f"❌ Data inválida: '{data_ref}'. Use o formato YYYYMMDD.")
    sys.exit(1)

print(f"📅 Data de referência: {data_ref}")

def ler_ficheiro(caminho: str) -> pd.DataFrame:
    if caminho.lower().endswith(".csv"):
        df_raw = pd.read_csv(caminho)
    else:
        df_raw = pd.read_excel(caminho)

    # O ficheiro deve ter 6 colunas (U_NAME,OPR_DATE,OPR_TIME,HOST_REF_NUMBER,CREATION_DATE_MOV,TASK_CODE)
    # O CONTRATO será obtido via JOIN com HUB.BRUCONTAS
    if df_raw.shape[1] < 6:
        raise ValueError("O ficheiro de Mobile Banking precisa de 6 colunas "
                          "(U_NAME,OPR_DATE,OPR_TIME,HOST_REF_NUMBER,CREATION_DATE_MOV,TASK_CODE).")

    df_raw = df_raw.iloc[:, :6]
    df_raw.columns = ["U_NAME", "OPR_DATE", "OPR_TIME", "HOST_REF_NUMBER", "CREATION_DATE_MOV", "TASK_CODE"]

    # Converter tipos de dados conforme especificado
    df_raw["U_NAME"] = df_raw["U_NAME"].astype(str).str.strip()
    df_raw["OPR_DATE"] = pd.to_datetime(df_raw["OPR_DATE"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    df_raw["OPR_TIME"] = pd.to_numeric(df_raw["OPR_TIME"], errors="coerce")
    df_raw["HOST_REF_NUMBER"] = pd.to_numeric(df_raw["HOST_REF_NUMBER"], errors="coerce")
    df_raw["CREATION_DATE_MOV"] = pd.to_numeric(df_raw["CREATION_DATE_MOV"], errors="coerce")
    df_raw["TASK_CODE"] = df_raw["TASK_CODE"].astype(str).str.strip()

    return df_raw

def criar_tabela_temporaria(cursor, df: pd.DataFrame, nome_tabela: str = "TEMP_MAPA7_MOBILE"):
    """Cria uma tabela temporária com os dados do DataFrame para uso na consulta SQL"""
    
    df_clean = df.copy()
    
    # Escapar aspas simples para evitar erro no SQL
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            df_clean[col] = df_clean[col].astype(str).str.replace("'", "''")
    
    # Criar a tabela temporária (sem CONTRATO, pois será obtido via JOIN)
    cursor.execute(f"""
        CREATE GLOBAL TEMPORARY TABLE {nome_tabela} (
            U_NAME VARCHAR2(100),
            OPR_DATE DATE,
            OPR_TIME NUMBER,
            HOST_REF_NUMBER NUMBER,
            CREATION_DATE_MOV NUMBER,
            TASK_CODE VARCHAR2(50)
        ) ON COMMIT PRESERVE ROWS
    """)
    
    # Inserir os dados do DataFrame
    for _, row in df_clean.iterrows():
        # Converter OPR_DATE para o formato Oracle
        if pd.notna(row['OPR_DATE']):
            if isinstance(row['OPR_DATE'], pd.Timestamp):
                opr_date_str = row['OPR_DATE'].strftime('%Y-%m-%d %H:%M:%S')
            else:
                opr_date_str = str(row['OPR_DATE'])
        else:
            opr_date_str = None
        
        # Preparar os valores para inserção
        u_name = str(row['U_NAME']) if pd.notna(row['U_NAME']) else None
        opr_time = float(row['OPR_TIME']) if pd.notna(row['OPR_TIME']) else None
        host_ref = float(row['HOST_REF_NUMBER']) if pd.notna(row['HOST_REF_NUMBER']) else None
        creation_date = float(row['CREATION_DATE_MOV']) if pd.notna(row['CREATION_DATE_MOV']) else None
        task_code = str(row['TASK_CODE']) if pd.notna(row['TASK_CODE']) else None
        
        # Inserir linha
        cursor.execute(f"""
            INSERT INTO {nome_tabela} (U_NAME, OPR_DATE, OPR_TIME, HOST_REF_NUMBER, CREATION_DATE_MOV, TASK_CODE)
            VALUES (:1, TO_DATE(:2, 'YYYY-MM-DD HH24:MI:SS'), :3, :4, :5, :6)
        """, (u_name, opr_date_str, opr_time, host_ref, creation_date, task_code))
    
    cursor.execute(f"COMMIT")
    print(f"✅ Tabela temporária {nome_tabela} criada com {len(df_clean)} registos.")

try:
    df_ficheiro = ler_ficheiro(ficheiro)
except Exception as e:
    print(f"❌ Erro ao ler o ficheiro: {e}")
    sys.exit(1)

if df_ficheiro.empty:
    print("⚠️ O ficheiro não tem linhas válidas. Nenhum ficheiro foi gerado.")
    sys.exit(0)

excel_path = "C:/Users/sum001781/Desktop/projeto/templates/Mapa 7a - Transacções efectuadas por clientes subscritores de mobile banking.xlsx"

# Query — usando a tabela temporária com JOIN a HUB.BRUCONTAS
query = f"""
WITH NUIB AS(
    SELECT ICNTIT, ICVCOP
    FROM (
        SELECT ICNTIT, ICVCOP, ICDTCR, 
               RANK() OVER(PARTITION BY ICNTIT ORDER BY ICDTCR DESC, dtref DESC, ICHOCR DESC) RK
        FROM HUB.IFCOPC
        WHERE ICCCOP = 'NUIB' AND ICSITU = 'V'
    )
    WHERE RK = 1
),
-- CTE que faz o JOIN com HUB.BRUCONTAS para obter o CONTRATO
MAPA7_MOBILE AS (
    SELECT 
        T.U_NAME,
        T.OPR_DATE,
        T.OPR_TIME,
        T.HOST_REF_NUMBER,
        T.CREATION_DATE_MOV,
        T.TASK_CODE,
        T1.BRCUNCLI AS CONTA,
        T1.BRCUCONTA AS CONTRATO
    FROM TEMP_MAPA7_MOBILE T
    LEFT JOIN HUB.BRUCONTAS T1 ON LOWER(T.U_NAME) = LOWER(T1.BRCUCODU)
)
SELECT /*+ PARALLEL */
    CASE WHEN COALESCE(T3.NUIB, T9.ICVCOP) IS NULL THEN 'N/A' ELSE COALESCE(T3.NUIB, T9.ICVCOP) END AS NUIB,
    CASE WHEN T4.N_OPERACAO IS NULL THEN 0 ELSE T4.N_OPERACAO END AS "ID da transacção",
    CASE WHEN T8.NIB IS NULL THEN 'N/A' ELSE T8.NIB END AS "NIB de débito",
    CASE WHEN T3.TEL_1 IS NULL THEN 'N/A' ELSE T3.TEL_1 END AS "Número de telemóvel de débito",
    T.OPR_DATE AS "Data e hora de operação",
    CASE 
        WHEN T.TASK_CODE IN ('PC_F_GCRNIFT','PC_F_CRNIFT','PC_F_CRNSFT') THEN '01'
        WHEN T.TASK_CODE IN ('PC_F_GCRNDFT','PC_F_CRNDFT') THEN '02'
        WHEN T.TASK_CODE = 'DW_T_BN' THEN '04'
        WHEN T.TASK_CODE IN ('CC_F_CCPC','EB_F_ISP','EB_F_JSP','LN_F_LRP','EB_F_PSP') THEN '05'
        WHEN T.TASK_CODE = 'EB_F_RO' THEN '06'
        WHEN T.TASK_CODE IN ('PC_F_GCRNDFTSI','PC_F_GCRNIFTSI','PC_F_CRNDFTSI','PC_F_CRNIFTSI') THEN '10'
        ELSE 'N/A' 
    END AS "Tipo de Operação",
    T4.MOEDA_AFECTA_CONTA AS "Moeda",
    NVL(TO_CHAR(ABS(MAX(CASE 
                        WHEN (T4.CODIGO_OPERACAO IN (40,542,580,602,604,640,726,46,432,50,434,56,180,108) 
                              OR (T4.CODIGO_OPERACAO = 260 AND T4.DESCRITIVO_MOV NOT LIKE '%OIC')
                              OR (T4.CODIGO_OPERACAO = 700 AND NVL(T4.GRUPO_PRECARIO, ' ') NOT IN ('DE', 'CO'))
                             )
                        THEN T4.VALOR_AFECTA_CONTA 
                    END 
                     ) 
                ) 
           ),'N/A'
        ) AS Valor,
    NVL(TO_CHAR(ABS(MAX(CASE 
                        WHEN (T4.CODIGO_OPERACAO IN (40,542,580,602,604,640,726,46,432,50,434,56,180,108) 
                              OR (T4.CODIGO_OPERACAO = 260 AND T4.DESCRITIVO_MOV NOT LIKE '%OIC')
                              OR (T4.CODIGO_OPERACAO = 700 AND NVL(T4.GRUPO_PRECARIO, ' ') NOT IN ('DE', 'CO'))
                             )
                        THEN T4.VALOR_AFECTA_CONTA 
                    END 
                     ) 
                ) 
           ),'N/A'
        ) AS "Valor na moeda original",
    NVL(TO_CHAR(ABS(MAX(CASE 
                                WHEN (T4.CODIGO_OPERACAO NOT IN (40,542,580,602,604,640,726,46,432,50,434,56,180,108) 
                                      OR (T4.CODIGO_OPERACAO = 260 AND T4.DESCRITIVO_MOV LIKE '%OIC')
                                      OR (T4.CODIGO_OPERACAO = 700 AND NVL(T4.GRUPO_PRECARIO, ' ') IN ('DE', 'CO'))
                                     )
                                THEN T4.VALOR_AFECTA_CONTA 
                            END ) ) ),
            'N/A'
        ) AS "Comissão",
    CASE WHEN T7.NIB IS NULL THEN 'N/A' ELSE T7.NIB END AS "NIB de crédito",
    CASE WHEN T7.TEL_1 IS NULL THEN 'N/A' ELSE T7.TEL_1 END AS "Número de telemóvel de crédito"
FROM MAPA7_MOBILE T
LEFT JOIN DWMB.FAC_MOVS_DO T4 ON T.CONTRATO = T4.CONTRATO AND 
                                (T.HOST_REF_NUMBER = T4.N_OPERACAO OR
                                  (T.HOST_REF_NUMBER IS NULL AND T4.N_OPERACAO IS NOT NULL AND T.CREATION_DATE_MOV = TO_NUMBER(TO_CHAR(T4.DATA_SISTEMA,'YYYYMMDD')) AND T4.HORA_SISTEMA = T.OPR_TIME) OR
                                  (T.HOST_REF_NUMBER IS NULL AND T4.N_OPERACAO IS NOT NULL AND T.CREATION_DATE_MOV = TO_NUMBER(TO_CHAR(T4.DATA_SISTEMA,'YYYYMMDD')) 
                                    AND TO_NUMBER(SUBSTR(TO_CHAR(T4.HORA_SISTEMA, '000000'), 1, 5)) = TO_NUMBER(SUBSTR(TO_CHAR(T.OPR_TIME, '000000'), 1, 5))
                                  )
                                )
LEFT JOIN DWMB.LKP_ENTIDADE T3 ON T4.COD_ENT = T3.COD_ENT
LEFT JOIN DWMB.LKP_OPR T6 ON T6.OPR = T4.CODIGO_OPERACAO
LEFT JOIN DWMB.T_DIM_CONTRATO T8 ON T8.COD_CONTRATO = T4.CONTRATO AND TO_DATE ({data_ref}, 'YYYYMMDD') BETWEEN T8.DT_INIC_REG AND T8.DT_FIM_REG
LEFT JOIN NUIB T9 ON T3.COD_ENT = T9.ICNTIT
LEFT JOIN (
    SELECT 
       T2.N_OPERACAO,
       T4.TEL_1,
       T5.NIB
    FROM DWMB.FAC_MOVS_DO T2
    LEFT JOIN DWMB.LKP_OPR T3 ON T3.OPR = T2.CODIGO_OPERACAO
    LEFT JOIN DWMB.LKP_ENTIDADE T4 ON T4.COD_ENT = T2.COD_ENT
    LEFT JOIN DWMB.T_DIM_CONTRATO T5 ON T5.COD_CONTRATO = T2.CONTRATO AND TO_DATE({data_ref}, 'YYYYMMDD') BETWEEN T5.DT_INIC_REG AND T5.DT_FIM_REG
    WHERE NVL(T2.PARAMETRO_ESTORNO_ANULACAO, ' ') NOT IN ('A', 'E')
      AND T2.GRUPO_PRECARIO NOT IN ('IM', 'JR')
      AND T3.TIPO_MOV = 'C'
      AND T2.CANAL IN ('INTERNET', 'MOBILE')
) T7 ON T7.N_OPERACAO = T4.N_OPERACAO 
WHERE T.CONTRATO = T4.CONTRATO 
AND (T.HOST_REF_NUMBER = T4.N_OPERACAO OR
      (T.HOST_REF_NUMBER IS NULL AND T4.N_OPERACAO IS NOT NULL AND T.CREATION_DATE_MOV = TO_NUMBER(TO_CHAR(T4.DATA_SISTEMA,'YYYYMMDD')) AND T4.HORA_SISTEMA = T.OPR_TIME) OR
      (T.HOST_REF_NUMBER IS NULL AND T4.N_OPERACAO IS NOT NULL AND T.CREATION_DATE_MOV = TO_NUMBER(TO_CHAR(T4.DATA_SISTEMA,'YYYYMMDD')) 
        AND TO_NUMBER(SUBSTR(TO_CHAR(T4.HORA_SISTEMA, '000000'), 1, 5)) = TO_NUMBER(SUBSTR(TO_CHAR(T.OPR_TIME, '000000'), 1, 5))
      )
    )
AND NVL(T4.PARAMETRO_ESTORNO_ANULACAO, ' ') NOT IN ('A', 'E')
AND T4.GRUPO_PRECARIO NOT IN ('IM', 'JR')
AND T6.TIPO_MOV = 'D'
AND T4.CANAL IN ('INTERNET', 'MOBILE')
GROUP BY
    COALESCE(T3.NUIB, T9.ICVCOP),
    T4.N_OPERACAO,
    T8.NIB,
    T3.TEL_1,
    T4.MOEDA_AFECTA_CONTA,
    T7.NIB,
    T7.TEL_1,
    T.TASK_CODE, 
    T.OPR_DATE
"""

# ── Executar query via pool ───────────────────
print("⏳ A executar consulta...")

try:
    with connection() as conn:
        cursor = conn.cursor()
        
        # 1. Criar tabela temporária com os dados do ficheiro
        nome_tabela_temp = "TEMP_MAPA7_MOBILE"
        
        # Verificar se a tabela já existe e dropar
        try:
            cursor.execute(f"DROP TABLE {nome_tabela_temp}")
            conn.commit()
        except oracledb.Error:
            pass  # Tabela não existe, continuar
        
        # Criar a tabela temporária e inserir os dados
        criar_tabela_temporaria(cursor, df_ficheiro, nome_tabela_temp)
        
        # 2. Executar a consulta usando a tabela temporária
        cursor.execute(query)
        columns = [col[0] for col in cursor.description]
        data = cursor.fetchall()
        df = pd.DataFrame(data, columns=columns)
        
        # 3. Limpar: dropar a tabela temporária
        try:
            cursor.execute(f"DROP TABLE {nome_tabela_temp}")
            conn.commit()
        except oracledb.Error:
            pass

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    # Tentar dropar tabela em caso de erro
    try:
        cursor.execute(f"DROP TABLE {nome_tabela_temp}")
        conn.commit()
    except:
        pass
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

# ── Formatação dos campos ───────────────────
if 'NUIB' in df.columns:
    df['NUIB'] = df['NUIB'].astype(str).str.zfill(15)
if 'NIB de débito' in df.columns:
    df['NIB de débito'] = df['NIB de débito'].astype(str).str.zfill(21)
if 'NIB de crédito' in df.columns:
    df['NIB de crédito'] = df['NIB de crédito'].astype(str).str.zfill(21)

# ── Sem dados: não gera ficheiro ──
if df.empty:
    print("⚠️ A consulta não retornou resultados. Nenhum ficheiro foi gerado.")
    sys.exit(0)

# ── Carregar template Excel ───────────────────
try:
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.worksheets[0]
    
    # REMOVER TODAS AS VALIDAÇÕES DE DADOS (DROPDOWNS)
    ws.data_validations = None
    
except Exception as e:
    print(f"❌ Erro ao carregar template Excel: {e}")
    sys.exit(1)

# ── Preencher cabeçalho ──────────────────────
data_ref_dt  = datetime.strptime(data_ref, "%Y%m%d")
primeiro_dia = datetime(data_ref_dt.year, data_ref_dt.month, 1)
ultimo_dia   = datetime(
    data_ref_dt.year,
    data_ref_dt.month,
    calendar.monthrange(data_ref_dt.year, data_ref_dt.month)[1]
)

ws["C4"] = str(data_ref_dt.year)
ws["C4"].number_format = "@"

ws["C5"].number_format = "@"
ws["C5"] = primeiro_dia.strftime('%Y-%m-%d')

ws["C6"].number_format = "@"
ws["C6"] = ultimo_dia.strftime('%Y-%m-%d')

# ── Preparar dados para Excel ─────────────────
# Ordem das colunas conforme o template (C → S)
colunas_template = [
    'NUIB',                            # C
    'ID da transacção',                # D
    'NIB de débito',                  # E
    'Número de telemóvel de débito',  # F
    'Data e hora de operação',        # G
    'Tipo de Operação',               # H
    'Moeda',                          # I
    'Valor',                          # J
    'Valor na moeda original',        # K
    'Comissão',                       # L
    'NIB de crédito',                 # M
    'Número de telemóvel de crédito'  # N
]

colunas_faltantes = [col for col in colunas_template if col not in df.columns]
if colunas_faltantes:
    print(f"⚠️ Colunas faltantes: {colunas_faltantes}")

df_template = df[colunas_template].copy()

# Formatar data e hora
df_template['Data e hora de operação'] = pd.to_datetime(df_template['Data e hora de operação']).dt.strftime('%d-%m-%Y %H:%M:%S')

# Colunas que devem ficar como texto (para não virarem notação científica) — índices 0-based
COLUNAS_TEXTO = {0, 1, 2, 10}  # NUIB, ID da transacção, NIB de débito, NIB de crédito

FONTE_NORMAL = Font(bold=False)  # reset explícito — não herda negrito do template

print("\n📝 A carregar dados no Excel...")
dados_para_excel = df_template.values

for i, linha in enumerate(dados_para_excel):
    linha_excel = i + 11  # Começar na linha 11
    for j, valor in enumerate(linha):
        coluna_excel = j + 1  # Começar na coluna A
        cell = ws.cell(row=linha_excel, column=coluna_excel)
        cell.value = valor
        cell.font = FONTE_NORMAL

        if j in COLUNAS_TEXTO:
            cell.value = str(valor)
            cell.number_format = "@"

print("✅ Dados inseridos!")

# Formatar colunas de montante como números (H, I, J = colunas 8, 9, 10 - 1-based)
for row in range(11, 11 + len(df_template)):
    for col in [8, 9, 10]:  # Colunas H, I, J
        cell = ws.cell(row=row, column=col)
        if cell.value == 'N/A':
            pass  # Não faz nada, mantém o valor como está
        elif cell.value is not None:
            try:
                # Tenta converter para número
                float(cell.value)
                cell.number_format = '#,##0.00'
            except (ValueError, TypeError):
                pass

# ── Guardar ficheiro ──────────────────────────
nome_base = os.path.splitext(os.path.basename(excel_path))[0]
nome_ficheiro = f"{nome_base}_{data_ref}.xlsx"
caminho_final = os.path.join(pasta_destino, nome_ficheiro)

try:
    wb.save(caminho_final)
    print(f"✅ Ficheiro guardado em: {caminho_final}")
except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)

print("✅ Processo concluído com sucesso!")