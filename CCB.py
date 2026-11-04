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
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM PREFIXO_CCB PASTA_DESTINO")
    sys.exit(1)

data_inicio   = sys.argv[1]
data_fim      = sys.argv[2]
prefixo_ccb   = sys.argv[3].strip()
pasta_destino = sys.argv[4]

try:
    datetime.strptime(data_inicio, "%Y%m%d")
    datetime.strptime(data_fim,    "%Y%m%d")
except ValueError:
    print(f"❌ Datas inválidas: '{data_inicio}' / '{data_fim}'. Use YYYYMMDD.")
    sys.exit(1)

if not prefixo_ccb.isdigit():
    print(f"❌ Prefixo de CCB inválido: '{prefixo_ccb}'. Deve conter apenas dígitos.")
    sys.exit(1)

print(f"📅 Período: {data_inicio} → {data_fim}")
print(f"🔎 Prefixo CCB: {prefixo_ccb}%")


# ── Query ─────────────────────────────────────
# Nota: o prefixo do CCB é passado como bind variable (:ccb_prefixo), nunca
# concatenado directamente na string SQL — assim evitamos SQL injection,
# já que este valor vem de um input do utilizador.
query = """
SELECT /*+PARALLEL*/ SK_DATA_INT,UNEG,CONTRATO,T4.COMPONENTE,T3.NOME,MOEDA, CCB,N_OPERACAO,USER_LANCAMENTO, DSC_MOV,SUM(VALOR) VALOR
FROM
(
    SELECT /*+PARALLEL*/ N_OPERACAO, SK_DATA_INT, VALOR_MOVIMENTO AS VALOR,CCB,MOEDA, UNIDADE_NEGOCIO AS UNEG,CONTRATO
    , UTILIZADOR_LANCAMENTO AS USER_LANCAMENTO, DESCRITIVO_MOVIMENTO AS DSC_MOV
    FROM DWMB.FAC_MOV_CONTAB_BANKA
    WHERE SK_DATA_INT BETWEEN :data_inicio AND :data_fim AND (CCB LIKE :ccb_prefixo) 

    UNION ALL 

    SELECT /*+PARALLEL*/ N_OPERACAO, TO_NUMBER(TO_CHAR(DIA_INTEGRACAO, 'YYYYMMDD')) AS SK_DATA_INT, VALOR_AFECTA_CONTA  AS VALOR,CONTA_CCB_MOV CCB,MOEDA_TRANSACAO MOEDA, UNIDADE_NEGOCIO AS UNEG,CONTRATO
    , UTILIZADOR_LANCAMENTO AS USER_LANCAMENTO, DESCRITIVO_MOV AS DSC_MOV
    FROM DWMB.FAC_MOVS_DO 
    WHERE TO_NUMBER(TO_CHAR(DIA_INTEGRACAO, 'YYYYMMDD')) BETWEEN :data_inicio AND :data_fim AND (CONTA_CCB_MOV LIKE :ccb_prefixo) 

    UNION ALL 

    SELECT /*+PARALLEL*/ N_OPERACAO, SK_DATA_INT AS SK_DATA_INT, VALOR_MOVIMENTO AS VALOR,CCB,MOEDA, UNIDADE_NEGOCIO AS UNEG,CONTRATO
    , UTILIZADOR_LANCAMENTO AS USER_LANCAMENTO, DESCRITIVO_MOVIMENTO AS DSC_MOV
    FROM DWMB.FAC_MOV_DP
    WHERE SK_DATA_INT BETWEEN :data_inicio AND :data_fim AND (CCB LIKE :ccb_prefixo) 

    UNION ALL 

    SELECT /*+PARALLEL*/ NUM_OPERACAO AS N_OPERACAO, SK_DATA_INTEGRACAO AS SK_DATA_INT, VALOR_MOVIMENTO AS VALOR,CCB,COD_MOEDA_OPER MOEDA, UNEG,CONTRATO
    , UTILIZADOR_INSERCAO AS USER_LANCAMENTO, DESCRICAO_MOVIMENTO AS DSC_MOV
    FROM DWMB.FAC_MOV_CCRED
    WHERE SK_DATA_INTEGRACAO BETWEEN :data_inicio AND :data_fim AND (CCB LIKE :ccb_prefixo) 

    UNION ALL 

    SELECT /*+PARALLEL*/ N_OPERACAO, SK_DATA_INT AS SK_DATA_INT, VALOR_MOVIMENTO AS VALOR,CCB,MOEDA, UNIDADE_NEGOCIO AS UNEG,CONTRATO
    , UTILIZADOR_LANCAMENTO AS USER_LANCAMENTO, DESCRITIVO_MOVIMENTO AS DSC_MOV
    FROM DWMB.FAC_MOV_CREDITO
    WHERE SK_DATA_INT BETWEEN :data_inicio AND :data_fim AND (CCB LIKE :ccb_prefixo) 

    UNION ALL 

    SELECT /*+PARALLEL*/ NUM_OPERACAO, SK_DATA_INTEGRACAO AS SK_DATA_INT, VALOR_MOVIMENTO AS VALOR, CCB_MOVIMENTO CCB,MOEDA, UNIDADE_NEGOCIO AS UNEG,CONTRATO
    , UTILIZADOR_INSERCAO AS USER_LANCAMENTO, DESCRICAO_MOVIMENTO AS DSC_MOV
    FROM DWMB.FAC_MOV_CC
    WHERE SK_DATA_INTEGRACAO BETWEEN :data_inicio AND :data_fim AND (CCB_MOVIMENTO LIKE :ccb_prefixo) 

    UNION ALL 

    SELECT /*+PARALLEL*/ DISTINCT N_OPERACAO, SK_DATA_INT AS SK_DATA_INT, VALOR_MOVIMENTO AS VALOR, CCB,MOEDA, UNIDADE_NEGOCIO AS UNEG,CONTRATO
    , UTILIZADOR_LANCAMENTO AS USER_LANCAMENTO, DESCRITIVO_MOVIMENTO AS DSC_MOV
    FROM DWMB.FAC_MOV_LEASE
    WHERE SK_DATA_INT BETWEEN :data_inicio AND :data_fim AND (CCB LIKE :ccb_prefixo) 

    UNION ALL

    SELECT /*+PARALLEL*/ N_OPERACAO, SK_DATA_INT AS SK_DATA_INT, VALOR_MOVIMENTO AS VALOR,CCB,MOEDA, UNIDADE_NEGOCIO AS UNEG,CONTRATO
    , UTILIZADOR_LANCAMENTO  AS USER_LANCAMENTO, DESCRITIVO_MOVIMENTO AS DSC_MOV
    FROM DWMB.FAC_MOV_EFT
    WHERE SK_DATA_INT BETWEEN :data_inicio AND :data_fim AND (CCB LIKE :ccb_prefixo) 

) T1
LEFT JOIN DWMB.LKP_CONTRATO T2 ON T2.COD_CONTRATO=T1.CONTRATO
LEFT JOIN DWMB.LKP_ENTIDADE T3 ON T3.COD_ENT=T2.COD_ENT
LEFT JOIN DWMB.DIM_PRODUTOS T4 ON T4.COMPONENTE=T2.COD_COMP_RNA
WHERE (CCB LIKE :ccb_prefixo)

GROUP BY SK_DATA_INT,UNEG,CONTRATO,T4.COMPONENTE,T3.NOME,MOEDA,N_OPERACAO, USER_LANCAMENTO, DSC_MOV, CCB
"""

bind_vars = {
    "data_inicio": int(data_inicio),
    "data_fim": int(data_fim),
    "ccb_prefixo": f"{prefixo_ccb}%",
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
        # Limpar apenas colunas de texto problemáticas
        colunas_texto = ['DSC_MOV']  # ajuste conforme necessário
        df = limpar_dataframe(df, colunas_especificas=colunas_texto)

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

# ── Gerar Excel ───────────────────────────────
nome_ficheiro = f"Extratos_CCB_{prefixo_ccb}_{data_inicio}_{data_fim}.xlsx"
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