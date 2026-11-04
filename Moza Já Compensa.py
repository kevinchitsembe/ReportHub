# QUERY QUE NÃO LEVA DATAS, SÓ FILTRO
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
if len(sys.argv) < 3:
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM FICHEIRO_OPERACOES PASTA_DESTINO")
    sys.exit(1)

ficheiro_entidades = sys.argv[1]
pasta_destino     = sys.argv[2]

# ── Ler ficheiro de operacoes ─────────────────
# header=True    → primeira linha é cabeçalho ("entidade")
# coluna=0       → primeira coluna
# numerico=True  → COD_ENT é numérico na Oracle, sem aspas no SQL
operacoes     = ler_coluna(ficheiro_entidades, coluna=0, header=True)
where_operacoes = formatar_in_clause(operacoes, numerico=True, coluna="T.N_OPERACAO")


# ── Query ─────────────────────────────────────
query = f"""
SELECT DISTINCT T.CONTRATO, T.N_OPERACAO, T2.CCB,  T2.UNIDADE_NEGOCIO UN_CCB, T.CONTA_CCB_MOV, T.UNIDADE_NEGOCIO UN_CCB_MOV
FROM dwmb.fac_movs_do T
LEFT JOIN dwmb.lkp_opr T1 ON T1.OPR = T.CODIGO_OPERACAO
LEFT JOIN (SELECT DISTINCT T.CONTRATO, T.N_OPERACAO, T.CCB, T.UNIDADE_NEGOCIO
           FROM DWMB.fac_mov_contab_banka T
           WHERE {where_operacoes} 
             AND T.CODIGO_OPERACAO IN (563, 583)
          ) T2 ON T2.CONTRATO = T.CONTRATO AND T2.N_OPERACAO = T.N_OPERACAO
WHERE T1.TIPO_MOV = 'D' AND {where_operacoes} 
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
nome_ficheiro = f"Moza_Já_Compensa.xlsx"
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