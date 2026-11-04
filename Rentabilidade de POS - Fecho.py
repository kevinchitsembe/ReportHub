"""
Fecho de POS.py
-----------------------
Fecho de POS filtrados por lista de entidades (ficheiro upload).

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
#print(f"📂 Ficheiro de contratos: {os.path.basename(ficheiro_entidades)}")


# ── Ler ficheiro de entidades ─────────────────
# header=True    → primeira linha é cabeçalho ("entidade")
# coluna=0       → primeira coluna
# numerico=True  → COD_ENT é numérico na Oracle, sem aspas no SQL
entidades     = ler_coluna(ficheiro_entidades, coluna=0, header=True)
where_contratos = formatar_in_clause(entidades, numerico=True, coluna="T.CONTRATO")
# ── Query ─────────────────────────────────────
query = f"""
SELECT *
FROM
(
    SELECT  T.DIA, T1.COD_ENT, T1.NOME, T.CONTRATO, T.UNIDADE_NEGOCIO, T.CODIGO_OPERACAO, T.DESCRITIVO_MOV,  
    CASE WHEN T.CODIGO_OPERACAO = 208 THEN -VALOR_AFECTA_CONTA ELSE VALOR_AFECTA_CONTA 
    END AS VALOR
    FROM DWMB.FAC_MOVS_DO T
    JOIN DWMB.LKP_ENTIDADE T1 ON T.COD_ENT = T1.COD_ENT
    WHERE T.CODIGO_OPERACAO IN (208, 209) AND 
        T.DIA IS NOT NULL AND 
        T.DIA BETWEEN {data_inicio} AND {data_fim} AND 
        {where_contratos}
) 
PIVOT (SUM(VALOR) FOR CODIGO_OPERACAO IN (208 COMISSÃO, 209 VALOR)
) 
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
        colunas_texto = ['DESCRITIVO_MOV']  # ajuste conforme necessário
        df = limpar_dataframe(df, colunas_especificas=colunas_texto)

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

#print("🔒 Conexão devolvida ao pool.")

# ── Gerar Excel ───────────────────────────────
nome_ficheiro = f"Fecho_POS_{data_inicio}_{data_fim}.xlsx"
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