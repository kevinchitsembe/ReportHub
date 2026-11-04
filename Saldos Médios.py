import sys
import os
import pandas as pd
import oracledb
from datetime import datetime

# Adicionar raiz do projecto ao path para importar utilitários partilhados
sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from upload_utils import ler_coluna, formatar_in_clause
from db import connection

# ── Argumentos ────────────────────────────────
if len(sys.argv) < 4:
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM FICHEIRO_ENTIDADES PASTA_DESTINO")
    sys.exit(1)

data_inicio       = sys.argv[1]
data_fim          = sys.argv[2]
pasta_destino     = sys.argv[3]

try:
    datetime.strptime(data_inicio, "%Y%m%d")
    datetime.strptime(data_fim,    "%Y%m%d")
except ValueError:
    print(f"❌ Datas inválidas: '{data_inicio}' / '{data_fim}'. Use YYYYMMDD.")
    sys.exit(1)

print(f"📅 Período: {data_inicio} → {data_fim}")
#print(f"📂 Ficheiro de contratos: {os.path.basename(ficheiro_entidades)}")

# ── Query ─────────────────────────────────────
query = f"""
SELECT /*+PARALLEL*/ COD_CNT, ROUND(ABS(SUM(VAL_SLD_CONV)) / COUNT(VAL_SLD_CONV), 2) SALDOS_MEDIOS
FROM DWMB.T_FAC_SLD_CNT
WHERE SK_TEMP_SLD_CNT BETWEEN {data_inicio} AND {data_fim} 
        AND COD_CLASSE_COMP_RNA IN ('DO', 'DP')
GROUP BY COD_CNT
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
        #print(f"📋 {len(df)} registo(s) encontrado(s).")

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

#print("🔒 Conexão devolvida ao pool.")

# ── Gerar Excel ───────────────────────────────
nome_ficheiro = f"Saldos_Médios_{data_inicio}_{data_fim}.xlsx"
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