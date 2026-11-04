import sys
import os
import pandas as pd
import oracledb
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font
import calendar
sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection

# ── Argumentos ────────────────────────────────
if len(sys.argv) < 4:
    print("❌ Uso: python script.py DATA_INICIO DATA_FIM PASTA_DESTINO")
    sys.exit(1)

data_inicio   = sys.argv[1]
data_fim      = sys.argv[2]
pasta_destino = sys.argv[3]

try:
    datetime.strptime(data_inicio, "%Y%m%d")
    datetime.strptime(data_fim,    "%Y%m%d")
except ValueError:
    print(f"❌ Datas inválidas: '{data_inicio}' / '{data_fim}'. Use YYYYMMDD.")
    sys.exit(1)

print(f"📅 Período: {data_inicio} → {data_fim}")

excel_path = "C:/Users/sum001781/Desktop/projeto/templates/Mapa 1b - Operações por Agência Bancária.xlsx"

sql_drop_quantidade = "DROP TABLE QUANTIDADE_CONTAS"
sql_drop_contratos = "DROP TABLE CONTRATOS_CREDITOS"
sql_drop_transacoes = "DROP TABLE TRANSACOES_bdm"

sql_create_quantidade = f"""
CREATE TABLE QUANTIDADE_CONTAS AS 
SELECT COD_UNID_NEG,DSC_UNIDADENEGOCIO, COUNT (COD_CONTRATO) QUANTIDADE_CONTAS_ABERTAS,  COUNT (CONTA_BASICA) QUANTIDADE_CONTAS_BASICAS
FROM (
    SELECT COD_CONTRATO, COD_COMP, COD_COMP_RNA, COD_CLASSE_COMP_RNA, COD_UNID_NEG,
    CASE WHEN cod_comp_rna  = 'DO_BAS23' then 1 ELSE NULL END CONTA_BASICA,DSC_UNIDADENEGOCIO
    FROM DWMB.T_DIM_CONTRATO T
    JOIN DWMB.DIM_ORGANIZACIONAL T1 ON T.COD_UNID_NEG = T1.UNIDADENEGOCIO
    WHERE DT_ABERT BETWEEN TO_DATE({data_inicio}, 'yyyymmdd') AND TO_DATE({data_fim}, 'yyyymmdd')
    AND TO_DATE({data_fim}, 'yyyymmdd') BETWEEN  DT_INIC_REG and DT_FIM_REG 
    AND COD_CLASSE_COMP = 'DO'
)
GROUP BY COD_UNID_NEG, DSC_UNIDADENEGOCIO
"""

sql_create_contratos = f"""
CREATE TABLE CONTRATOS_CREDITOS AS 
SELECT COD_UNID_NEG, COUNT (COD_CONTRATO) QUANTIDADE_CONTAS_CREDITO,  SUM(val_contrato_CONV) VALOR_CREDITOS
FROM (
SELECT t.cod_unid_neg, T.COD_ENT,T.COD_CNT,T.COD_CONTRATO,T.DT_ABERT,T.DT_VENC,T2.NATUREZA_2,T2.DSC_NATUREZA_2,
T2.TIPO_PRODUTO,T2.DSC_TIPO_PRODUTO, val_contrato, (t.val_contrato * TX_COMPRA) val_contrato_CONV
FROM DWMB.T_DIM_CONTRATO T
JOIN DWMB.DIM_PRODUTOS T2 ON T2.COMPONENTE=T.COD_COMP_RNA AND T2.NATUREZA_2='DES'
LEFT JOIN DWMB.T_FAC_CAMBIO T3 ON T.DT_ABERT= t3.DT_COT AND t3.cod_moeda = t.COD_MOEDA AND t3.cod_moeda_base = 'MZN' AND t3.cod_tipo_cot = 'R'
WHERE  TO_DATE({data_fim}, 'yyyymmdd') BETWEEN  DT_INIC_REG and DT_FIM_REG and DT_ABERT BETWEEN TO_DATE({data_inicio}, 'yyyymmdd') AND TO_DATE({data_fim}, 'yyyymmdd')
)
GROUP BY COD_UNID_NEG
"""

sql_create_transacoes = f"""
CREATE TABLE TRANSACOES_bdm AS 
select UNIDADE_NEGOCIO, count (n_operacao_dePOSITO_METICAIS) "Quantidade de depósitos em meticais",
sum (valor_transacao_DEPOSITO) "Valor de Depósitos ",
count (n_operacao_dePOSITO_ME) "Quantidade de depósitos em moeda estrangeira",
sum (valor_transacao_DEPOSITO_ME) "Valor de Depósitos (ME)",
count (n_operacao_levantamento) "Quantidade de Levantamentos",
sum (valor_levantamento) "Valor de Levantamentos ",
count (n_operacao_TRANSF_CRD_MZN) "Quantidade de Transferências em moeda nacional crédito meticais",
sum (VALOR_TRANSF_CRD_MZN) "Valor de Transferências credito meticais",
count (n_operacao_TRANSF_DEB_MZN) "Quantidade de Transferências em moeda nacional débito meticais",
sum (VALOR_TRANSF_DEB_MZN) "Valor de Transferências débito meticais",
count (n_operacao_TRANSF_CRD_ESTRANG) "Quantidade de Transferências em moeda estrangeira crédito",
sum (VALOR_TRANSF_CRD_ESTRANG) "Valor de Transferências (ME) credito",
count (n_operacao_TRANSF_DEB_ESTRANG) "Quantidade de Transferências em moeda estrangeira débito",
sum (VALOR_TRANSF_DEB_ESTRANG) "Valor de Transferências (ME) débito"

from(


SELECT/*+parallel*/
    t.dia, 
    t.contrato,
    UNIDADE_NEGOCIO,
    -- depósios em meticais 
    CASE WHEN TIPO_MOV = 'C' AND CODIGO_OPERACAO IN (1,5,13,60,61,491,493,841,843) and moeda_transacao = 'MZN' THEN T.n_operacao END AS n_operacao_dePOSITO_METICAIS,
     CASE WHEN TIPO_MOV = 'C' AND CODIGO_OPERACAO IN (1,5,13,60,61,491,493,841,843) and moeda_transacao = 'MZN' THEN (T.valor_transacao * t2.TX_COMPRA) ELSE 0 END AS valor_transacao_DEPOSITO,
     
    -- depósios em meticais 
    CASE WHEN TIPO_MOV = 'C' AND CODIGO_OPERACAO IN (1,5,13,60,61,491,493,841,843) and moeda_transacao <> 'MZN' THEN T.n_operacao END AS n_operacao_dePOSITO_ME,
     CASE WHEN TIPO_MOV = 'C' AND CODIGO_OPERACAO IN (1,5,13,60,61,491,493,841,843) and moeda_transacao <> 'MZN' THEN (T.valor_transacao * t2.TX_COMPRA) ELSE 0 END AS valor_transacao_DEPOSITO_ME,

    -- Levantamentos não especifica se é ATM, não especifica se é pelo balcão de domicilio ou movimento
    CASE WHEN TIPO_MOV = 'D' AND (CODIGO_OPERACAO IN (2,4,6,7,8,9,12,14,32,214,274,610) OR UPPER(DESCRITIVO_MOV) LIKE '%LEVANTAMENT%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN T.n_operacao END AS n_operacao_levantamento,
    CASE WHEN TIPO_MOV = 'D' AND (CODIGO_OPERACAO IN (2,4,6,7,8,9,12,14,32,214,274,610) OR UPPER(DESCRITIVO_MOV) LIKE '%LEVANTAMENT%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN (T.valor_transacao * t2.TX_COMPRA) END AS valor_levantamento,     
    -- Transferências Crédito em meticais
    CASE WHEN TIPO_MOV = 'C' AND MOEDA_TRANSACAO ='MZN' AND( CODIGO_OPERACAO <> '725'  OR codigo_operacao in (63, 41, 441, 189, 183, 45, 47, 471, 467, 615, 211, 411, 221) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN T.n_operacao END AS n_operacao_TRANSF_CRD_MZN,
    CASE WHEN TIPO_MOV = 'C' AND MOEDA_TRANSACAO ='MZN' AND( CODIGO_OPERACAO <> '725'  OR codigo_operacao in (63, 41, 441, 189, 183, 45, 47, 471, 467, 615, 211, 411, 221) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN (T.valor_transacao * t2.TX_COMPRA) END AS VALOR_TRANSF_CRD_MZN,
    -- Transferências Débito em meticais
    CASE WHEN TIPO_MOV = 'D' AND MOEDA_TRANSACAO ='MZN' AND (CODIGO_OPERACAO <> '730' OR codigo_operacao in (220, 210, 410, 614, 476, 46, 44, 182, 150, 188, 40, 62) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN T.n_operacao END AS n_operacao_TRANSF_DEB_MZN,
     CASE WHEN TIPO_MOV = 'D' AND MOEDA_TRANSACAO ='MZN' AND (CODIGO_OPERACAO <> '730' OR codigo_operacao in (220, 210, 410, 614, 476, 46, 44, 182, 150, 188, 40, 62) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN (T.valor_transacao * t2.TX_COMPRA) END AS VALOR_TRANSF_DEB_MZN,
          -- Transferências Crédito em moeda estrangeira
    CASE WHEN TIPO_MOV = 'C' AND MOEDA_TRANSACAO <> 'MZN' AND( CODIGO_OPERACAO <> '725'  OR codigo_operacao in (63, 41, 441, 189, 183, 45, 47, 471, 467, 615, 211, 411, 221) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN T.n_operacao END AS n_operacao_TRANSF_CRD_ESTRANG,
    CASE WHEN TIPO_MOV = 'C' AND MOEDA_TRANSACAO <> 'MZN'  AND( CODIGO_OPERACAO <> '725'  OR codigo_operacao in (63, 41, 441, 189, 183, 45, 47, 471, 467, 615, 211, 411, 221) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN (T.valor_transacao * t2.TX_COMPRA) END AS VALOR_TRANSF_CRD_ESTRANG,
    -- Transferências Débito em moeda estrangeira
    CASE WHEN TIPO_MOV = 'D' AND MOEDA_TRANSACAO <> 'MZN' AND (CODIGO_OPERACAO <> '730' OR codigo_operacao in (220, 210, 410, 614, 476, 46, 44, 182, 150, 188, 40, 62) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN T.n_operacao END AS n_operacao_TRANSF_DEB_ESTRANG,
     CASE WHEN TIPO_MOV = 'D' AND MOEDA_TRANSACAO <> 'MZN' AND (CODIGO_OPERACAO <> '730' OR codigo_operacao in (220, 210, 410, 614, 476, 46, 44, 182, 150, 188, 40, 62) OR UPPER(DESCRITIVO_MOV) LIKE 'TRF%' OR UPPER(DESCRITIVO_MOV) LIKE 'TRANSFER%') 
         AND NVL(GRUPO_PRECARIO, ' ') NOT IN ('CO', 'DE','IMP') THEN (T.valor_transacao * t2.TX_COMPRA) END AS VALOR_TRANSF_DEB_ESTRANG 
                
FROM dwmb.fac_movs_do t
--JOIN dwmb.T_DIM_ENT T0 ON T.COD_ENT = T0.COD_ENT AND &DTRF_FIM between T0.DT_INIC_REG and T0.DT_FIM_REG
-- Cambio opcional: Fazemos o Join com a tabela de cambio
LEFT JOIN dwmb.t_fac_cambio t2 ON t.dia = t2.sk_temp_cot AND t2.cod_moeda = t.moeda_transacao AND t2.cod_moeda_base = 'MZN' AND t2.cod_tipo_cot = 'R'
join dwmb.lkp_opr t1 on t.codigo_operacao = opr
WHERE t.DIA BETWEEN {data_inicio} AND {data_fim}
AND NVL(parametro_estorno_anulacao, ' ') NOT IN ('A','E')
 and
CODIGO_OPERACAO <> '725'
)

group by UNIDADE_NEGOCIO
"""

sql_select = """
select t.COD_UNID_NEG "ID do Ponto",
NVL (QUANTIDADE_CONTAS_ABERTAS, 0) "Quantidade de Contas Abertas",
NVL (QUANTIDADE_CONTAS_BASICAS, 0) "Quantidade de Contas Básicas", 
NVL ("Quantidade de depósitos em meticais",0) "Quantidade de depósitos em meticais",
round  (Abs (NVL ("Valor de Depósitos ", 0)), 1) "Valor de Depósitos",
NVL ("Quantidade de depósitos em moeda estrangeira",0) "Quantidade de depósitos em moeda estrangeira",
round (Abs (NVL ("Valor de Depósitos (ME)", 0)), 1) "Valor de Depósitos (ME)",
NVL ("Quantidade de Levantamentos", 0) "Quantidade de Levantamentos",
round (Abs ( NVL ("Valor de Levantamentos ", 0)), 1) "Valor de Levantamentos",
NVL (QUANTIDADE_CONTAS_CREDITO, 0) "Quantidade de Créditos",
round (Abs (NVL (VALOR_CREDITOS, 0)), 1) "Valor de Créditos",
NVL ("Quantidade de Transferências em moeda nacional débito meticais", 0) "Quantidade de Transferências em moeda nacional",
round (Abs (NVL ("Valor de Transferências débito meticais", 0)), 1) "Valor de Transferências",
NVL ("Quantidade de Transferências em moeda estrangeira débito",0) "Quantidade de Transferências em moeda estrangeira",
round (Abs (NVL ("Valor de Transferências (ME) débito", 0)), 1) "Valor de Transferências (ME)"
from QUANTIDADE_CONTAS t
left join TRANSACOES_bdm t1 on t.COD_UNID_NEG = t1.UNIDADE_NEGOCIO
LEFT JOIN CONTRATOS_CREDITOS T2 ON  t.COD_UNID_NEG = t2.COD_UNID_NEG
"""
# ── Executar via pool ─────────────────────────
print("⏳ A executar consulta...")

try:
    with connection() as conn:
        cursor = conn.cursor()

        # 1. DROPs — ignorar erro se a tabela não existir
        for sql_drop in [sql_drop_quantidade, sql_drop_contratos, sql_drop_transacoes]:
            try:
                cursor.execute(sql_drop)
                conn.commit()
            except oracledb.Error as e:
                if "ORA-00942" in str(e):
                    print(f"A continuar...")
                else:
                    raise

        # 2. CREATEs
        cursor.execute(sql_create_quantidade)
        conn.commit()
        #print("🏗️ Tabela QUANTIDADE_CONTAS criada.")

        cursor.execute(sql_create_contratos)
        conn.commit()
        #print("🏗️ Tabela CONTRATOS_CREDITOS criada.")

        cursor.execute(sql_create_transacoes)
        conn.commit()
        #print("🏗️ Tabela TRANSACOES_bdm criada.")

        # 3. SELECT final
        cursor.execute(sql_select)
        columns = [col[0] for col in cursor.description]
        data = cursor.fetchall()
        df = pd.DataFrame(data, columns=columns)
        #print(f"📋 {len(df)} registo(s) encontrado(s).")

        # 4. DROPs finais — limpar tabelas temporárias
        for sql_drop in [sql_drop_quantidade, sql_drop_contratos, sql_drop_transacoes]:
            try:
                cursor.execute(sql_drop)
                conn.commit()
            except Exception:
                pass

except oracledb.Error as e:
    print(f"❌ Erro Oracle: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Erro ao executar query: {e}")
    sys.exit(1)

# ── Sem dados ─────────────────────────────────
if df.empty:
    print("⚠️ A consulta não retornou resultados. Nenhum ficheiro foi gerado.")
    sys.exit(0)

try:
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.worksheets[0]
    ws.data_validations = None
except Exception as e:
    print(f"❌ Erro ao carregar template Excel: {e}")
    sys.exit(1)

# ── Preencher cabeçalho ──────────────────────
data_ref_dt  = datetime.strptime(data_inicio, "%Y%m%d") 
primeiro_dia = datetime(data_ref_dt.year, data_ref_dt.month, 1)
ultimo_dia   = datetime(
    data_ref_dt.year,
    data_ref_dt.month,
    calendar.monthrange(data_ref_dt.year, data_ref_dt.month)[1]
)

ws["C4"] = str(data_ref_dt.year)
ws["C4"].number_format = "@"
# Converter data_inicio para datetime e formatar
data_inicio_dt = datetime.strptime(data_inicio, "%Y%m%d")
ws["C5"].number_format = "@"
ws["C5"] = data_inicio_dt.strftime('%d/%m/%Y')

# Converter data_fim para datetime e formatar (se existir)
if data_fim:
    data_fim_dt = datetime.strptime(data_fim, "%Y%m%d")
    ws["C6"].number_format = "@"
    ws["C6"] = data_fim_dt.strftime('%d/%m/%Y')
else:
    ws["C6"].number_format = "@"
    ws["C6"] = data_inicio_dt.strftime('%d/%m/%Y')  # ou deixar vazio

print("📝 A carregar dados no Excel...")

colunas_template = [
    'ID do Ponto',                                          # C (j=0)
    'Quantidade de Contas Abertas',                         # D (j=1)
    'Quantidade de Contas Básicas',                         # E (j=2)
    'Quantidade de depósitos em meticais',                  # F (j=3)
    'Valor de Depósitos',                                   # G (j=4)
    'Quantidade de depósitos em moeda estrangeira',         # H (j=5)
    'Valor de Depósitos (ME)',                              # I (j=6)
    'Quantidade de Levantamentos',                          # J (j=7)
    'Valor de Levantamentos',                               # K (j=8)
    'Quantidade de Créditos',                               # L (j=9)
    'Valor de Créditos',                                    # M (j=10)
    'Quantidade de Transferências em moeda nacional',       # N (j=11)
    'Valor de Transferências',                              # O (j=12)
    'Quantidade de Transferências em moeda estrangeira',    # P (j=13)
    'Valor de Transferências (ME)'                          # Q (j=14)
]

colunas_faltantes = [col for col in colunas_template if col not in df.columns]
if colunas_faltantes:
    print(f"⚠️ Colunas faltantes: {colunas_faltantes}")

df_template = df[colunas_template].copy()

# Colunas por tipo de formatação (índices 0-based, tal como em colunas_template)
COLUNAS_INTEIRO   = {0, 1, 2, 3, 5, 7, 9, 11, 13}   # IDs e "Quantidade de..."
COLUNAS_MONTANTE  = {4, 6, 8, 10, 12, 14}           # "Valor de..."

FONTE_NORMAL = Font(bold=False)  # reset explícito — não herda negrito do template

dados_para_excel = df_template.values

for i, linha in enumerate(dados_para_excel):
    linha_excel = i + 11
    for j, valor in enumerate(linha):
        coluna_excel = j + 1
        cell = ws.cell(row=linha_excel, column=coluna_excel)
        cell.value = valor
        cell.font = FONTE_NORMAL  # limpa qualquer negrito residual do template

        if j in COLUNAS_INTEIRO:
            cell.number_format = '###0'
        elif j in COLUNAS_MONTANTE:
            cell.number_format = '#,##0.00'

print("✅ Dados inseridos!")

nome_base = os.path.splitext(os.path.basename(excel_path))[0]
nome_ficheiro = f"{nome_base}_{data_inicio}_{data_fim}.xlsx"
caminho_final = os.path.join(pasta_destino, nome_ficheiro)

try:
    wb.save(caminho_final)
except Exception as e:
    print(f"❌ Erro ao guardar ficheiro Excel: {e}")
    sys.exit(1)

#print(f"✅ Ficheiro guardado em: {caminho_final}")