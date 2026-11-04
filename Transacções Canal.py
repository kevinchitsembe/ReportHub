import sys
import os
import pandas as pd
import oracledb
from datetime import datetime

sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection
from upload_utils import limpar_dataframe
# ── Argumentos ────────────────────────────────
# modo simples: script.py DATA_REFERENCIA PASTA_DESTINO filtro_tipo filtro_valor
if len(sys.argv) < 4:
    print("❌ Uso: python script.py DATA_REFERENCIA PASTA_DESTINO")
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

# ── Query ─────────────────────────────────────
query = f"""
select /*+ PARALLEL */ DISTINCT T.DIA,T.DATA_SISTEMA, T.HORA_SISTEMA, T.CONTRATO,
t1.nome,T1.EMAIL, T1.TEL_1, T1.TEL_2, T1.TEL_3, T8.DSC_SEG,
t2.dsc_segmentocomercial, t.codigo_operacao, T.n_operacao, T9.DSC_OPR DESCRITIVO_OPR, T.canal, 
T.valor_transacao*T4.tx_compra valor_transacao_mzn,
t2.DSC_UNIDADENEGOCIO, T.DESCRITIVO_MOV, NVL(t3.TIPO_TRANSF,' ') TIPO_TRANSF, t.estacao_lancamento
from dwmb.fac_movs_do T
left join dwmb.lkp_entidade t1 on t1.cod_ent=t.cod_ent 
LEFT JOIN DWMB.T_FAC_CAMBIO T4 ON T4.SK_TEMP_COT=T.DIA AND T.MOEDA_TRANSACAO = T4.COD_MOEDA AND T4.COD_TIPO_COT = 'R' AND T4.COD_MOEDA_BASE ='MZN'
LEFT JOIN DWMB.DIM_ORGANIZACIONAL t2 ON t2.UNIDADENEGOCIO=T.UNIDADE_NEGOCIO
left join dwmb.FAC_TEI_ENV t3 on T3.N_OPERACAO = T.N_OPERACAO
JOIN DWMB.LKP_OPR T9 ON T9.OPR = T.CODIGO_OPERACAO
LEFT JOIN DWMB.T_DIM_CNT t8 ON T8.SK_CNT=T.SK_CNT
where T.data_sistema  between to_date({data_inicio}, 'yyyymmdd') and to_date({data_fim} , 'YYYYMMDD') 
and NVL(T.PARAMETRO_ESTORNO_ANULACAO,' ') NOT IN('A','E') 
and (T.canal in ('WHATSAPP','USSD','INTERNET') OR T.UTILIZADOR_LANCAMENTO='INTUSR01') AND T.valor_transacao < 0 and t2.FLAG_BALCAO='SIM' AND t8.COD_SEG <> 'PT'
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
        # Limpar apenas colunas de texto problemáticas
        colunas_texto = ['DESCRITIVO_MOV']  # ajuste conforme necessário
        df = limpar_dataframe(df, colunas_especificas=colunas_texto)

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

# ── Gerar Excel na pasta temp/ ────────────────
nome_ficheiro = f"Transacções_Canal_{data_inicio}_{data_fim}.xlsx"
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

   # print(f"✅ Ficheiro gerado: {nome_ficheiro}")

except Exception as e:
    print(f"❌ Erro ao gerar ficheiro Excel: {e}")
    sys.exit(1)