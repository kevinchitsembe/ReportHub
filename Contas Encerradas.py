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
select /*+ PARALLEL */ a11.COD_CNT  COD_CNT,
a11.COD_CONTRATO  COD_CONTRATO,
a12.COD_ENT,
TO_CHAR(a11.DT_ABERT, 'YYYY-MM-DD') DT_ABERT,
a11.DT_VENC  DT_VENC,
a11.COD_SIT  COD_SIT,
a11.COD_CLASSE_COMP  COD_CLASSE_COMP,
a15.COD_GESTOR,
a15.cod_gestor_rna,
a10.nome NOME_GESTOR,
a12.NOME  NOME,
a15.COD_SEG  COD_SEG,
a15.DSC_SEG  DSC_SEG,
a14.NATUREZA_1  COD_NAT_1,
a14.DSC_NATUREZA_1  DSC_NAT_1,
a14.TIPO_PRODUTO  COD_TIPO_PROD,
a12.TEL_1  TEL_1,
a12.TEL_2  TEL_2,
a12.TEL_3  TEL_3,
a12.EMAIL  EMAIL,
a11.DSC_MOTIVO_SIT  DSC_MOTIVO_SIT,
TO_CHAR(a11.DT_SIT, 'YYYY-MM-DD') DT_SIT,
a11.COD_UNID_NEG  COD_UNID_NEG
from dwmb.lkp_contrato a11
join dwmb.dim_produtos t3 on a11.cod_comp_rna = t3.componente
join dwmb.lkp_entidade a12 on a12.cod_ent = a11.cod_ent
join dwmb.lkp_conta a15 on a15.cod_cnt = a11.cod_cnt
left join dwmb.dim_organizacional t6 on a15.COD_UNID_NEG = t6.unidadenegocio
join DWMB.DIM_PRODUTOS a14 on (a11.COD_COMP_RNA = a14.COMPONENTE)
join DWMB.lkp_gestor a10 on a10.cod_gestor = a15.cod_gestor_rna
where a11.COD_CLASSE_COMP in ('DO') and a11.cod_sit='E' and a11.dt_sit between to_date({data_inicio}, 'YYYYMMDD') and to_date({data_fim}, 'YYYYMMDD')
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
        colunas_texto = ['DSC_MOTIVO_SIT']  # ajuste conforme necessário
        df = limpar_dataframe(df, colunas_especificas=colunas_texto)

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

# ── Gerar Excel na pasta temp/ ────────────────
nome_ficheiro = f"Contas_Encerradas_{data_inicio}_{data_fim}.xlsx"
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