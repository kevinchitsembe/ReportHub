import sys
import os
import pandas as pd
import oracledb
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, numbers
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(__import__('pathlib').Path(__file__).parent.parent))
from db import connection

# ── Argumentos ────────────────────────────────
if len(sys.argv) < 3:
    print("❌ Uso: python mapa_9.py DATA_REF PASTA_DESTINO")
    print("   Exemplo: python mapa_9.py 20260701 C:/destino")
    sys.exit(1)

data_ref = sys.argv[1]
pasta_destino = sys.argv[2]

try:
    data_dt = datetime.strptime(data_ref, "%Y%m%d")
except ValueError:
    print(f"❌ Data inválida: '{data_ref}'. Use YYYYMMDD.")
    sys.exit(1)

print(f"📅 Data de referência: {data_ref}")

# ── Calcular datas ────────────────────────────
dt_inicio = data_ref
dt_fim = (data_dt.replace(day=28) + pd.Timedelta(days=4)).replace(day=1) - pd.Timedelta(days=1)
dt_fim_str = dt_fim.strftime('%Y%m%d')
ult_dia_mes_ant = (data_dt.replace(day=1) - pd.Timedelta(days=1)).strftime('%Y%m%d')

print(f"📅 Período: {dt_inicio} a {dt_fim_str}")
print(f"📅 Último dia mês anterior: {ult_dia_mes_ant}")

template_path = "C:/Users/sum001781/Desktop/projeto/templates/Mapa 9-Reporte das Novas Operações de Empréstimos realizados pelo Banco.xlsx"

# ── Queries separadas ─────────────────────────
sql_drop = "DROP TABLE DETALHE"

sql_create = f"""
CREATE TABLE DETALHE AS 
    WITH SLD_DESC_N_AUTORIZ AS
    (
      SELECT *
      FROM
      (
          SELECT  T1.SK_CNT,T1.COD_CNT, T1.SK_ENT, T1.COD_ENT, T1.SK_CONTRATO, T1.COD_CONTRATO, T1.COD_COMP_RNA, T1.COD_CLASSE_COMP_RNA, 
          T1.DT_SLD_CNT AS DT_ABERT, T1.DT_SLD_CNT + 1 AS DT_VENC, T1.VAL_TX_NOMINAL, T1.SPREAD, T1.VAL_TX_REF, T1.VAL_SLD_DT_VALOR AS MONTANTE_MT,
          LAG(VAL_SLD_DT_VALOR, 1, 0) OVER (PARTITION BY T1.COD_CONTRATO ORDER BY DT_SLD_CNT) AS SLD_DT_ANT,T3.COD_MOEDA
          FROM DWMB.T_FAC_SLD_CNT T1 
          INNER JOIN DWMB.T_DIM_CONTRATO T3
          ON T1.SK_CONTRATO=T3.SK_CONTRATO
           WHERE T1.COD_CLASSE_COMP='DO' AND SK_TEMP_SLD_CNT BETWEEN {ult_dia_mes_ant} AND TO_NUMBER(TO_CHAR(LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD')), 'YYYYMMDD')) AND T3.VAL_LIM_DESC = 0 
      ) T
      WHERE T.MONTANTE_MT < 0 AND T.SLD_DT_ANT >=0 AND T.DT_ABERT BETWEEN TO_DATE('{data_ref}','YYYYMMDD') AND LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD')) 
    ),
    DESC_AUTORIZ  AS
    (
        SELECT DISTINCT COD_CONTRATO, DT_INIC_LIM_DESC, TO_NUMBER(TO_CHAR(DT_INIC_LIM_DESC, 'YYYYMMDD')) AS SK_TEMP_SLD_CNT
        FROM DWMB.T_FAC_SLD_CNT
        WHERE COD_CLASSE_COMP='DO' AND SK_TEMP_SLD_CNT BETWEEN {data_ref} AND TO_NUMBER(TO_CHAR(LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD')), 'YYYYMMDD')) AND DT_INIC_LIM_DESC IS NOT NULL
        AND DT_INIC_LIM_DESC BETWEEN TO_DATE('{data_ref}','YYYYMMDD') AND LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD')) 
    ),
    MATURIDADES AS
    (
      SELECT 1 AS DIAS_MIN, 7 AS DIAS_MAX, '7 DIAS' AS MAT FROM DUAL
      UNION ALL
      SELECT 8 AS DIAS_MIN, 15 AS DIAS_MAX, '15 DIAS' AS MAT FROM DUAL
      UNION ALL
      SELECT 16 AS DIAS_MIN, 30 AS DIAS_MAX, '30 DIAS' AS MAT FROM DUAL
      UNION ALL
      SELECT 31 AS DIAS_MIN, 45 AS DIAS_MAX, '45 DIAS' AS MAT FROM DUAL
      UNION ALL
      SELECT 46 AS DIAS_MIN, 60 AS DIAS_MAX, '60 DIAS' AS MAT FROM DUAL
      UNION ALL
      SELECT 61 AS DIAS_MIN, 90 AS DIAS_MAX, '90 DIAS' AS MAT FROM DUAL
      UNION ALL
      SELECT 91 AS DIAS_MIN, 180 AS DIAS_MAX, '180 DIAS' AS MAT FROM DUAL
      UNION ALL
      SELECT 181 AS DIAS_MIN, 365 AS DIAS_MAX, '1 ANO' AS MAT FROM DUAL
      UNION ALL
      SELECT 366 AS DIAS_MIN, 730 AS DIAS_MAX, '2 ANOS' AS MAT FROM DUAL
      UNION ALL
      SELECT 731 AS DIAS_MIN, 999999 AS DIAS_MAX, '+2 ANOS' AS MAT FROM DUAL
    ),
    DESCOBERTOS AS
    (
      SELECT 
        T.SK_CNT, 
        T.SK_ENT, 
        T.COD_ENT,
        T.COD_CNT,
        T.SK_CONTRATO, 
        T.COD_CONTRATO, 
        T.COD_COMP_RNA,
        T.COD_CLASSE_COMP_RNA, 
        T.DT_ABERT, 
        T.DT_VENC,
        T.VAL_TX_NOMINAL, 
        T.SPREAD, 
        T.VAL_TX_REF,
        T.MONTANTE_MT,T.COD_MOEDA
      FROM SLD_DESC_N_AUTORIZ  T
      WHERE T.MONTANTE_MT < 0 AND T.SLD_DT_ANT >=0 AND T.DT_ABERT BETWEEN TO_DATE('{data_ref}','YYYYMMDD') AND LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD'))
      UNION ALL
      SELECT 
        T.SK_CNT, 
        T.SK_ENT, 
        T.COD_ENT,
        T.COD_CNT,
        T.SK_CONTRATO, 
        T.COD_CONTRATO, 
        T.COD_COMP_RNA,
        T.COD_CLASSE_COMP_RNA, 
        T.DT_ABERT, 
        T.DT_VENC, 
        T.VAL_TX_NOMINAL, 
        T.SPREAD, 
        T.VAL_TX_REF,
        T.MONTANTE_MT,
        T.COD_MOEDA
      FROM
      (
        SELECT T1.SK_CNT,T1.COD_CNT, T1.SK_ENT, T1.COD_ENT, T1.SK_CONTRATO, T1.COD_CONTRATO, T1.COD_COMP_RNA, T1.COD_CLASSE_COMP_RNA, T1.DT_INIC_LIM_DESC AS DT_ABERT, T3.DT_EXP_LIM_DESC AS DT_VENC, T1.VAL_TX_NOMINAL, T1.SPREAD, T1.VAL_TX_REF, T1.VAL_CRD_CONV AS MONTANTE_MT,T1.COD_MOEDA
        FROM DWMB.T_FAC_SLD_CNT T1 INNER JOIN DESC_AUTORIZ T2
        ON T1.COD_CONTRATO=T2.COD_CONTRATO AND T1.SK_TEMP_SLD_CNT=T2.SK_TEMP_SLD_CNT
        INNER JOIN DWMB.T_DIM_CONTRATO T3
        ON T1.SK_CONTRATO=T3.SK_CONTRATO
      ) T
    )
SELECT  T2.NUIT, T2.NOME, TO_CHAR(T1.DT_ABERT,'DD-MM-YYYY') DT_ABERT,
        T2.DSC_EMP_PART AS TIPO_CLIENTE,
        ABS(T1.MONTANTE_MT) AS MONTANTE_MT,
        CASE T1.SPREAD 
          WHEN 0 THEN NULL
          ELSE T1.VAL_TX_NOMINAL - T1.SPREAD
        END AS PR, T1.SPREAD AS "SPREAD %", 
        ROUND(T1.VAL_TX_NOMINAL / 100, 2) AS "TAXA %",
        CASE T1.SPREAD 
          WHEN 0 THEN 'FIXA'
          ELSE 'VARIÁVEL'
        END AS TIPO_TAXA,
        MAT.MAT AS MATURIDADE,
        CASE T2.COD_EMP_PART 
          WHEN 'E' THEN 'APOIO À TESOURARIA' 
          ELSE 'CONSUMO' 
        END AS FINALIDADE,
        'DESCOBERTO BANCÁRIO' AS TIPO_CREDITO,
        T2.COD_CAE_2, T2.DSC_CAE_2,
        T1.COD_CLASSE_COMP_RNA, 
        DT_VENC - T1.DT_ABERT AS MATURIDADE_DIAS, 
        NULL AS COD_LINHA_CRED, T5.COD_SEG, T5.DSC_SEG, T1.COD_CONTRATO, TO_CHAR(DT_VENC,'DD-MM-YYYY') DT_VENC,
        TIPO_PRODUTO, DSC_TIPO_PRODUTO,T1.COD_MOEDA 
FROM DESCOBERTOS  T1
LEFT JOIN DWMB.T_DIM_ENT T2 ON T1.SK_ENT = T2.SK_ENT
LEFT JOIN DWMB.T_DIM_CNT T5 ON T1.SK_CNT=T5.SK_CNT
INNER JOIN DWMB.DIM_PRODUTOS T3 ON T1.COD_COMP_RNA = T3.COMPONENTE
LEFT JOIN MATURIDADES MAT ON T1.DT_VENC - T1.DT_ABERT BETWEEN MAT.DIAS_MIN AND MAT.DIAS_MAX
UNION ALL
SELECT  T2.NUIT, T2.NOME, 
        CASE WHEN TIPO_PRODUTO = 'CCC' THEN TO_CHAR(T1.DT_INIC_PREM,'DD-MM-YYYY')
            ELSE TO_CHAR(T1.DT_ABERT,'DD-MM-YYYY') 
        END AS DT_ABERT,
        T2.DSC_EMP_PART AS TIPO_CLIENTE,
        ABS(T4.VAL_CRD_CONV) AS MONTANTE_MT,
        CASE SPREAD 
          WHEN 0 THEN NULL
          ELSE T4.VAL_TX_NOMINAL - T4.SPREAD
        END AS PR, T4.SPREAD AS "SPREAD %", 
        ROUND(T4.VAL_TX_NOMINAL / 100, 2) AS "TAXA %",
        CASE SPREAD 
          WHEN 0 THEN 'FIXA'
          ELSE 'VARIÁVEL'
        END AS TIPO_TAXA,
        MAT.MAT AS MATURIDADE,
        CASE  
          WHEN TIPO_PRODUTO = 'CC' THEN 'CONSUMO'
          WHEN TIPO_PRODUTO = 'CCC' THEN 'APOIO À TESOURARIA'
          WHEN TIPO_PRODUTO = 'CON' THEN 'CONSUMO'
          WHEN TIPO_PRODUTO = 'INV' THEN 'INVESTIMENTO'
          WHEN TIPO_PRODUTO = 'CF_PUB_MZN' THEN 'CONSUMO'
          WHEN TIPO_PRODUTO = 'HAB' THEN 'HABITAÇÃO'
          WHEN TIPO_PRODUTO = 'LIV' THEN 'APOIO À TESOURARIA'
          WHEN TIPO_PRODUTO IN ('LEA', 'LES', 'LEI', 'LEM')  THEN 'LEASING'
        END AS FINALIDADE,
        CASE 
          WHEN T5.COD_SEG = 'PT' THEN 'CRÉDITO SOCIAL AO COLABORADOR'
          WHEN NVL(TRIM(T1.COD_LINHA_CRED), ' ') NOT IN (' ', '-1') THEN 'LINHA ESPECIAL'
          WHEN T1.COD_CLASSE_COMP_RNA = 'CARC' THEN 'CARTÃO DE CRÉDITO'
          ELSE 'NORMAL'
        END AS TIPO_CREDITO,
        T2.COD_CAE_2, T2.DSC_CAE_2,
        T1.COD_CLASSE_COMP_RNA, 
        NVL(T1.DT_VENC, T1.DT_ABERT+ 730) - T1.DT_ABERT AS MATURIDADE_DIAS, 
        T1.COD_LINHA_CRED, T5.COD_SEG, T5.DSC_SEG, T1.COD_CONTRATO, TO_CHAR(T1.DT_VENC,'DD-MM-YYYY') DT_VENC,
        TIPO_PRODUTO, DSC_TIPO_PRODUTO,T1.COD_MOEDA
FROM DWMB.T_DIM_CONTRATO T1
LEFT JOIN DWMB.T_DIM_ENT T2 ON T1.SK_ENT = T2.SK_ENT 
LEFT JOIN DWMB.T_DIM_CNT T5 ON T1.SK_CNT=T5.SK_CNT
INNER JOIN DWMB.DIM_PRODUTOS T3 ON T1.COD_COMP_RNA = T3.COMPONENTE
LEFT JOIN DWMB.T_FAC_SLD_CNT T4 ON T1.COD_CONTRATO=T4.COD_CONTRATO AND T4.SK_TEMP_SLD_CNT = TO_NUMBER(TO_CHAR(T1.DT_ABERT, 'YYYYMMDD'))
LEFT JOIN MATURIDADES MAT ON NVL(T1.DT_VENC, T1.DT_ABERT+ 730) - T1.DT_ABERT BETWEEN MAT.DIAS_MIN AND MAT.DIAS_MAX
LEFT JOIN 
    (
        SELECT COD_CONTRATO,DT_INIC_PREM, MIN(DT_INIC_REG) DT_INIC_REG  
        FROM DWMB.T_DIM_CONTRATO 
        WHERE COD_CLASSE_COMP_RNA='CC' AND DT_INIC_PREM BETWEEN TO_DATE('{data_ref}','YYYYMMDD') AND LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD')) AND DT_INIC_PREM<=DT_INIC_REG GROUP BY COD_CONTRATO,DT_INIC_PREM 
    ) A ON A.COD_CONTRATO=T1.COD_CONTRATO AND A.DT_INIC_REG=T1.DT_INIC_REG
WHERE (
        (T1.DT_ABERT BETWEEN TO_DATE('{data_ref}','YYYYMMDD') AND LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD')) AND T1.DT_ABERT BETWEEN T1.DT_INIC_REG AND T1.DT_FIM_REG) OR
        (T1.DT_INIC_PREM BETWEEN TO_DATE('{data_ref}','YYYYMMDD') AND LAST_DAY(TO_DATE('{data_ref}','YYYYMMDD')) AND A.DT_INIC_REG BETWEEN T1.DT_INIC_REG AND T1.DT_FIM_REG)
      ) 
      AND NOT (T1.COD_SIT ='E' AND T1.DT_SIT = T1.DT_ABERT)
      AND T3.CLASSE <> 'DS' AND T3.NATUREZA_2='DES'
"""

# 🔹 SELECT com ORDER BY consistente para garantir ROWNUM estável
sql_select = """
SELECT ROWNUM AS NUM_ORDEM, 
       T.NUIT, 
       T2.NOME, 
       T.DT_ABERT, 
       T.TIPO_CLIENTE, 
       T.MONTANTE_MT, 
       T.PR, 
       T."SPREAD %", 
       T."TAXA %", 
       T.TIPO_TAXA, 
       T.MATURIDADE, 
       T.FINALIDADE, 
       T.TIPO_CREDITO, 
       T.DSC_CAE_2
FROM (
    SELECT T.*
    FROM DETALHE T
    JOIN DWMB.LKP_CONTRATO T1 ON T.COD_CONTRATO=T1.COD_CONTRATO
    JOIN DWMB.LKP_ENTIDADE T2 ON T2.COD_ENT=T1.COD_ENT
    ORDER BY T.COD_CONTRATO, T.DT_ABERT, T.NUIT
) T
LEFT JOIN DWMB.LKP_CONTRATO T1 ON T.COD_CONTRATO=T1.COD_CONTRATO
LEFT JOIN DWMB.LKP_ENTIDADE T2 ON T2.COD_ENT=T1.COD_ENT
"""

# ── Executar via pool ─────────────────────────
print("⏳ A executar consulta...")

try:
    with connection() as conn:
        cursor = conn.cursor()

        # 1. DROP — ignorar erro se a tabela não existir
        try:
            cursor.execute(sql_drop)
            conn.commit()
        except oracledb.Error as e:
            if "ORA-00942" in str(e):
                print("ℹ️ Tabela DETALHE não existia, a continuar.")
            else:
                raise

        # 2. CREATE
        cursor.execute(sql_create)
        conn.commit()
        print("🏗️ Tabela DETALHE criada.")

        # 3. SELECT com ORDER BY consistente
        cursor.execute(sql_select)
        columns = [col[0] for col in cursor.description]
        data    = cursor.fetchall()
        df      = pd.DataFrame(data, columns=columns)
        print(f"📋 {len(df)} registo(s) encontrado(s).")

        # 4. DROP final — limpar tabela
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

# ── Abrir template e preencher ────────────────
print("📝 A carregar dados no Excel...")

try:
    wb = load_workbook(template_path)
    ws = wb.worksheets[0]

    # ── Preencher cabeçalho ──
    ws["C7"] = data_dt.year
    ws["C8"] = data_dt.strftime('%d/%m/%Y')
    ws["C9"] = dt_fim.strftime('%d/%m/%Y')
    
    for cell in ["C7", "C8", "C9"]:
        ws[cell].alignment = Alignment(horizontal="left")

    # ── Ordem das colunas (A → N) ──
    colunas_template = [
        'NUM_ORDEM', 'NUIT', 'NOME', 'DT_ABERT', 'TIPO_CLIENTE',
        'MONTANTE_MT', 'PR', 'SPREAD %', 'TAXA %', 'TIPO_TAXA',
        'MATURIDADE', 'FINALIDADE', 'TIPO_CREDITO', 'DSC_CAE_2'
    ]

    df_template = df[colunas_template].copy()
    
    # 🔹 Formatação NUIT: remover zeros à esquerda
    df_template['NUIT'] = df_template['NUIT'].astype(str).str.lstrip('0')
    df_template['NUIT'] = df_template['NUIT'].apply(lambda x: x if x else '0')
    
    # Outras formatações
    df_template['NUM_ORDEM'] = df_template['NUM_ORDEM'].astype(int)
    df_template['DT_ABERT'] = pd.to_datetime(df_template['DT_ABERT'], format='%d-%m-%Y', errors='coerce').dt.strftime('%d/%m/%Y')

    # Colunas de texto (índices 0-based)
    colunas_texto = {0, 1, 2, 3, 4, 9, 10, 11, 12, 13}

    # Inserir dados a partir da linha 13
    for i, linha in enumerate(df_template.values):
        linha_excel = i + 13
        for j, valor in enumerate(linha):
            coluna_excel = j + 1
            cell = ws.cell(row=linha_excel, column=coluna_excel)
            
            if j in colunas_texto:
                cell.value = str(valor) if pd.notna(valor) else ""
                cell.number_format = "@"
            else:
                if pd.notna(valor):
                    cell.value = valor
                    if coluna_excel == 6:  # MONTANTE_MT
                        cell.number_format = '#,##0.00'
                    elif coluna_excel in [7, 8, 9]:  # PR, SPREAD%, TAXA%
                        cell.number_format = '0.00%'
                else:
                    cell.value = None

    print("✅ Dados inseridos no template!")

    # ── Calcular totais ──
    soma_montante = df_template['MONTANTE_MT'].sum()
    ws["F7"] = soma_montante
    ws["F7"].number_format = '#,##0.00'
    
    max_taxa = df_template['TAXA %'].max()
    ws["F8"] = max_taxa if pd.notna(max_taxa) else 0
    ws["F8"].number_format = '0.00%'
    
    min_taxa = df_template['TAXA %'].min()
    ws["F9"] = min_taxa if pd.notna(min_taxa) else 0
    ws["F9"].number_format = '0.00%'
    
    linha_fim = 13 + len(df_template) - 1
    ws["F10"] = f'=IFERROR(SUMPRODUCT(F13:F{linha_fim},I13:I{linha_fim})/SUM(F13:F{linha_fim}),0)'
    ws["F10"].number_format = '0.00%'
    
    for cell in ["F7", "F8", "F9", "F10"]:
        ws[cell].alignment = Alignment(horizontal="left")

    print("✅ Totais calculados e preenchidos!")

    # ── Ajustar largura das colunas ──
    for col in range(1, 15):
        col_letter = get_column_letter(col)
        max_length = 0
        for row in range(1, 13 + len(df_template) + 5):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value is not None:
                cell_length = len(str(cell_value))
                max_length = max(max_length, cell_length)
        ws.column_dimensions[col_letter].width = min(max(max_length + 5, 12), 50)

    # ── Guardar ficheiro ──
    nome_base = os.path.splitext(os.path.basename(template_path))[0]
    nome_ficheiro = f"{nome_base}_{data_ref}.xlsx"
    caminho_final = os.path.join(pasta_destino, nome_ficheiro)
    
    # Criar pasta se não existir
    os.makedirs(pasta_destino, exist_ok=True)
    
    wb.save(caminho_final)
    print(f"✅ Ficheiro guardado em: {caminho_final}")
    print(f"📊 Total de registos: {len(df_template)}")
    print(f"💰 Soma MONTANTE_MT: {soma_montante:,.2f}")
    
    if pd.notna(max_taxa):
        print(f"📈 MAX TAXA %: {max_taxa:.2%}")
    if pd.notna(min_taxa):
        print(f"📉 MIN TAXA %: {min_taxa:.2%}")

except Exception as e:
    print(f"❌ Erro ao processar Excel: {e}")
    sys.exit(1)