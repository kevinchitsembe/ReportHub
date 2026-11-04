import sys
import os
import pandas as pd
import oracledb
from datetime import datetime

sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection

# ── Argumentos ────────────────────────────────
if len(sys.argv) < 5:
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM FICHEIRO_SALDOS_MEDIOS PASTA_DESTINO")
    sys.exit(1)

data_inicio        = sys.argv[1]
data_fim            = sys.argv[2]
ficheiro_saldos     = sys.argv[3]
pasta_destino       = sys.argv[4]

try:
    datetime.strptime(data_inicio, "%Y%m%d")
    datetime.strptime(data_fim,    "%Y%m%d")
except ValueError:
    print(f"❌ Datas inválidas: '{data_inicio}' / '{data_fim}'. Use YYYYMMDD.")
    sys.exit(1)

print(f"📅 Período: {data_inicio} → {data_fim}")


# ── Ler ficheiro de saldos médios ─────────────
# Assume que o ficheiro tem cabeçalho e que as duas primeiras colunas
# são, por ordem: código da conta e saldo médio (independentemente
# do nome exacto do cabeçalho no ficheiro do utilizador).
def ler_saldos_medios(caminho: str) -> pd.DataFrame:
    if caminho.lower().endswith(".csv"):
        df_raw = pd.read_csv(caminho)
    else:
        df_raw = pd.read_excel(caminho)

    if df_raw.shape[1] < 2:
        raise ValueError("O ficheiro de saldos médios precisa de pelo menos 2 colunas "
                          "(código da conta e saldo médio).")

    df_raw = df_raw.iloc[:, :2]
    df_raw.columns = ["COD_CNT", "SALDOS_MEDIOS"]

    # Normalizar tipos para o merge com a Oracle correr sem surpresas
    df_raw["COD_CNT"] = (
        df_raw["COD_CNT"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    )
    df_raw["SALDOS_MEDIOS"] = pd.to_numeric(df_raw["SALDOS_MEDIOS"], errors="coerce")
    df_raw = df_raw.dropna(subset=["SALDOS_MEDIOS"])

    return df_raw


try:
    df_saldos = ler_saldos_medios(ficheiro_saldos)
except Exception as e:
    print(f"❌ Erro ao ler o ficheiro de saldos médios: {e}")
    sys.exit(1)

if df_saldos.empty:
    print("⚠️ O ficheiro de saldos médios não tem linhas válidas. Nenhum ficheiro foi gerado.")
    sys.exit(0)


# ── Query Oracle (sem a tabela SALDOSMEDIOS — essa só existe no ficheiro) ──
# Nota: datas passadas como bind variables, não interpoladas na string SQL.
query = """
WITH MOVS_IB_USSD AS (
    SELECT /*+PARALLEL*/ DISTINCT T.COD_CNT, T.DESCRITIVO_MOV DESCRICAO_MOVIMENTO
    FROM DWMB.FAC_MOVS_DO T
    LEFT JOIN DWMB.LKP_OPR T1 ON T.CODIGO_OPERACAO = T1.OPR
    WHERE T.DIA BETWEEN :data_inicio AND :data_fim
        AND T.CANAL IN ('INTERNET','USSD')
        AND NVL(T.PARAMETRO_ESTORNO_ANULACAO,' ') NOT IN ('A','E')
        AND T.GRUPO_PRECARIO NOT IN ('CO','IM','DE','JR')
        AND T1.TIPO_MOV = 'D'
),
MOVS_CARTOES AS (
    SELECT /*+ PARALLEL */ DISTINCT
            A.COD_CNT, A.DESCRITIVO_MOV DESCRICAO_MOVIMENTO
    FROM DWMB.FAC_MOVS_DO A
    LEFT JOIN DWMB.DIM_CARTAO E ON SUBSTR(E.PAN,1,15) = SUBSTR(A.NUM_CARTAO,1,15)
    LEFT JOIN DWMB.LKP_OPR B ON B.OPR = A.CODIGO_OPERACAO
    WHERE A.SUBSISTEMA IS NOT NULL
        AND NVL(A.GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE', 'IM')
        AND NVL(A.PARAMETRO_ESTORNO_ANULACAO, ' ') NOT IN ('A', 'E')
        AND E.COMPONENTE NOT LIKE '%PP%'
        AND A.DIA BETWEEN :data_inicio AND :data_fim
        AND B.TIPO_MOV = 'D'
    UNION ALL
    SELECT /*+ PARALLEL */ DISTINCT
            A.CONTA COD_CNT, A.DESCRICAO_MOVIMENTO
    FROM DWMB.FAC_MOV_CCRED A
    LEFT JOIN DWMB.LKP_OPR B ON B.OPR = A.CODIGO_OPERACAO
    WHERE A.SUBSISTEMA IS NOT NULL
        AND NVL(A.GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE', 'IM')
        AND A.SK_DATA_INTEGRACAO BETWEEN :data_inicio AND :data_fim
        AND B.TIPO_MOV = 'D'
)
SELECT /*+ PARALLEL */ DISTINCT T.COD_CNT, T4.NOME
FROM MOVS_IB_USSD T
INNER JOIN MOVS_CARTOES T1 ON T.COD_CNT = T1.COD_CNT
LEFT JOIN DWMB.LKP_CONTA T3 ON T.COD_CNT = T3.COD_CNT
LEFT JOIN DWMB.LKP_ENTIDADE T4 ON T3.COD_ENT = T4.COD_ENT
WHERE T3.COD_SEG <> 'PT'
"""

bind_vars = {
    "data_inicio": int(data_inicio),
    "data_fim": int(data_fim),
}

print("⏳ A executar consulta...")

try:
    with connection() as conn:
        cursor = conn.cursor()
        cursor.execute(query, bind_vars)
        columns = [col[0] for col in cursor.description]
        data    = cursor.fetchall()
        df_oracle = pd.DataFrame(data, columns=columns)

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

# ── Cruzar com os saldos médios do ficheiro (equivalente ao INNER JOIN) ──
df_oracle["COD_CNT"] = df_oracle["COD_CNT"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

df = df_oracle.merge(df_saldos, on="COD_CNT", how="inner")
df = df[df["SALDOS_MEDIOS"] >= 1000]

# ── Gerar Excel ───────────────────────────────
if df.empty:
    print("⚠️ A consulta não devolveu dados para este período/ficheiro. Nenhum ficheiro foi gerado.")
    sys.exit(0)

df = df.fillna("N/A")

nome_ficheiro = f"Clientes_M_{data_inicio}_{data_fim}.xlsx"
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

except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)