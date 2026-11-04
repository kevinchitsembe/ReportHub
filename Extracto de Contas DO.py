import sys
import os
import pandas as pd
import oracledb
from datetime import datetime

sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection
from upload_utils import limpar_dataframe
# ── Argumentos ────────────────────────────────
if len(sys.argv) < 5:
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM COD_ENT PASTA_DESTINO")
    sys.exit(1)

data_inicio   = sys.argv[1]
data_fim      = sys.argv[2]
cod_ent       = sys.argv[3].strip()  # Agora é COD_ENT em vez de prefixo_ccb
pasta_destino = sys.argv[4]

try:
    datetime.strptime(data_inicio, "%Y%m%d")
    datetime.strptime(data_fim,    "%Y%m%d")
except ValueError:
    print(f"❌ Datas inválidas: '{data_inicio}' / '{data_fim}'. Use YYYYMMDD.")
    sys.exit(1)

# Validar se COD_ENT é numérico
if not cod_ent.isdigit():
    print(f"❌ COD_ENT inválido: '{cod_ent}'. Deve conter apenas dígitos.")
    sys.exit(1)

cod_ent_int = int(cod_ent)

print(f"📅 Período: {data_inicio} → {data_fim}")
print(f"🔎 COD_ENT: {cod_ent} (filtro exato)")

# ── Query ─────────────────────────────────────
# Nota: COD_ENT é passado como bind variable (:cod_ent), nunca concatenado
# diretamente na string SQL — assim evitamos SQL injection,
# já que este valor vem de um input do utilizador.
query = """
SELECT 
   T1.DIA, 
   T2.COD_ENT,
   T1.UNIDADE_NEGOCIO,
   T3.DSC_UNIDADENEGOCIO,
   T6.NOME,
   T2.COD_CNT, 
   T1.CONTRATO, 
   T1.N_DOCUMENTO, 
   T1.N_OPERACAO, 
   T5.DSC_OPR,
   T1.DESCRITIVO_MOV,
   T1.UTILIZADOR_LANCAMENTO,
   T1.MOEDA_TRANSACAO,
   T1.VALOR_TRANSACAO AS SALDO_NA_MOEDA_TRANSAÇÃO,
   CASE
       WHEN T1.MOEDA_TRANSACAO <> 'MZN' AND T7.TX_COMPRA IS NOT NULL THEN
           T1.VALOR_TRANSACAO * T7.TX_COMPRA 
       ELSE NULL
   END AS SALDO_CONVERTIDO
FROM 
   DWMB.FAC_MOVS_DO T1
INNER JOIN DWMB.LKP_CONTRATO T2
    ON T1.CONTRATO = T2.COD_CONTRATO
INNER JOIN DWMB.DIM_ORGANIZACIONAL T3
    ON T1.UNIDADE_NEGOCIO = T3.UNIDADENEGOCIO
INNER JOIN DWMB.LKP_ENTIDADE T4
    ON T2.COD_ENT = T4.COD_ENT
INNER JOIN DWMB.LKP_OPR T5
    ON T1.CODIGO_OPERACAO = T5.OPR
INNER JOIN DWMB.LKP_ENTIDADE T6
    ON T2.COD_ENT = T6.COD_ENT
LEFT JOIN DWMB.T_FAC_CAMBIO T7
    ON T7.COD_MOEDA = T1.MOEDA_TRANSACAO 
    AND T7.COD_TIPO_COT = 'R' 
    AND T7.SK_TEMP_COT = T1.DIA
WHERE 
    T2.COD_ENT = :cod_ent 
    AND T1.DIA BETWEEN :data_inicio AND :data_fim
"""

bind_vars = {
    "data_inicio": int(data_inicio),
    "data_fim": int(data_fim),
    "cod_ent": cod_ent_int,  # Filtro exato, sem LIKE
}

# ── Executar query via pool ───────────────────
print("⏳ A executar consulta...")

try:
    with connection() as conn:
        cursor = conn.cursor()
        cursor.execute(query, bind_vars)
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

# ── Gerar Excel ───────────────────────────────
# Nome do arquivo atualizado para refletir que é COD_ENT em vez de CCB
nome_ficheiro = f"Extrato_DO_{cod_ent}_{data_inicio}_{data_fim}.xlsx"
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

    print(f"✅ Ficheiro guardado: {caminho_final}")

except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)