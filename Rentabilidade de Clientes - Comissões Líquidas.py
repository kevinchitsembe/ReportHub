"""
mov_contab_entidades.py
-----------------------
Movimentos contabilísticos filtrados por lista de entidades (ficheiro upload).

sys.argv[1] → data_inicio  (YYYYMMDD)
sys.argv[2] → data_fim     (YYYYMMDD)
sys.argv[3] → caminho do ficheiro de entidades (xlsx ou csv) — enviado pela app
sys.argv[4] → pasta de destino (temp/)
"""

import sys
import os
import pandas as pd
import oracledb
from datetime import datetime

# Adicionar raiz do projecto ao path para importar utilitários partilhados
sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from upload_utils import ler_coluna, formatar_in_clause, limpar_dataframe
from db import connection

# ── Argumentos ────────────────────────────────
if len(sys.argv) < 5:
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM FICHEIRO_ENTIDADES PASTA_DESTINO")
    sys.exit(1)

data_inicio       = sys.argv[1]
data_fim          = sys.argv[2]
ficheiro_entidades = sys.argv[3]
pasta_destino     = sys.argv[4]

try:
    datetime.strptime(data_inicio, "%Y%m%d")
    datetime.strptime(data_fim,    "%Y%m%d")
except ValueError:
    print(f"❌ Datas inválidas: '{data_inicio}' / '{data_fim}'. Use YYYYMMDD.")
    sys.exit(1)

print(f"📅 Período: {data_inicio} → {data_fim}")
#print(f"📂 Ficheiro de entidades: {os.path.basename(ficheiro_entidades)}")


# ── Ler ficheiro de entidades ─────────────────
# header=True    → primeira linha é cabeçalho ("entidade")
# coluna=0       → primeira coluna
# numerico=True  → COD_ENT é numérico na Oracle, sem aspas no SQL
entidades     = ler_coluna(ficheiro_entidades, coluna=0, header=True)
entidades_sql = formatar_in_clause(entidades, numerico=True, coluna="T4.COD_ENT")


# ── Query ─────────────────────────────────────
query = f"""
SELECT /*+PARALLEL*/ DISTINCT T3.COD_CNT,T4.COD_ENT,T4.NOME, T1.CONTRATO, T1.SK_DATA_INT, T1.MOEDA, T5.TX_COMPRA CAMBIO_REAVALIACAO,
(CASE WHEN T2.PLN_NVL_3 = 2 AND T1.CONTRATO != 0 THEN T1.VALOR_MOVIMENTO * T5.TX_COMPRA ELSE 0 END) AS COM_LIQ_DIRECT_MZN, T1.DESCRITIVO_MOVIMENTO
FROM DWMB.FAC_MOV_CONTAB_BANKA T1 
LEFT JOIN DWMB.DIM_CONTAB T2 ON T1.CCB = T2.CCB
INNER JOIN DWMB.LKP_CONTRATO T3 ON T3.COD_CONTRATO=T1.CONTRATO
LEFT JOIN DWMB.LKP_ENTIDADE T4 ON T3.COD_ENT=T4.COD_ENT
LEFT JOIN DWMB.T_DIM_HIER_TEMP T5 ON T5.SK_TEMP=T1.SK_DATA_INT 
LEFT JOIN DWMB.T_FAC_CAMBIO T5 ON T5.SK_TEMP_COT=T1.SK_DATA_INT AND T1.MOEDA=T5.COD_MOEDA AND T5.COD_TIPO_COT = 'R' AND T5.COD_MOEDA_BASE = 'MZN'
WHERE T2.PLN_NVL_1=1 AND T2.PLN_NVL_3 IN (2, 1, 41) AND T1.SK_DATA_INT BETWEEN {data_inicio} AND {data_fim}
AND {entidades_sql} AND NVL(T1.PARAM_ESTORNO_ANULACAO,' ') NOT IN('A','E')
"""


# ── Executar query via pool ───────────────────
print("⏳ A executar consulta...")

try:
    with connection() as conn:
        #print("✅ Conexão com Oracle estabelecida com sucesso!")
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [col[0] for col in cursor.description]
        data    = cursor.fetchall()
        df      = pd.DataFrame(data, columns=columns)
        colunas_texto = ['DESCRITIVO_MOVIMENTO']  # ajuste conforme necessário
        df = limpar_dataframe(df, colunas_especificas=colunas_texto)

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

#print("🔒 Conexão devolvida ao pool.")

# ── Gerar Excel ───────────────────────────────
dados_vazios = df.empty
if dados_vazios:
    print("⚠️ A consulta não retornou resultados.")
    #df = pd.DataFrame([{col: "N/A" for col in columns}])

nome_ficheiro = f"Comissões_Líquidas_{data_inicio}_{data_fim}.xlsx"
caminho_final = os.path.join(pasta_destino, nome_ficheiro)

try:
    with pd.ExcelWriter(caminho_final, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Dados", index=False)

        ws = writer.sheets["Dados"]
        for idx, col in enumerate(df.columns):
            try:
                largura = max(df[col].astype(str).map(len).max(), len(col))
            except Exception:
                largura = len(col)
            letra = (
                chr(65 + idx)
                if idx < 26
                else chr(64 + idx // 26) + chr(65 + idx % 26)
            )
            ws.column_dimensions[letra].width = min(largura + 2, 50)

    #print(f"✅ Ficheiro guardado em: {caminho_final}")

except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)