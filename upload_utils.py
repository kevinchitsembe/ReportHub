"""
upload_utils.py
---------------
Utilitários reutilizáveis para leitura de ficheiros carregados pelo utilizador.
Importar em cada script que precise de ficheiro externo:

    from upload_utils import ler_coluna, ler_dataframe, formatar_in_clause, limpar_dataframe
"""

import os
import sys
import pandas as pd
import re  # ← ADICIONADO: necessário para regex
import unicodedata  # ← ADICIONADO: necessário para normalização Unicode


def ler_dataframe(caminho: str, header: bool = True) -> pd.DataFrame:
    """
    Lê um ficheiro .xlsx, .xls ou .csv e devolve um DataFrame.

    Parâmetros
    ----------
    caminho : str
        Caminho absoluto do ficheiro
    header : bool
        True  → primeira linha é cabeçalho (default)
        False → sem cabeçalho, colunas ficam 0, 1, 2...
    """
    ext = os.path.splitext(caminho)[1].lower()
    header_row = 0 if header else None

    try:
        if ext in [".xlsx", ".xls"]:
            return pd.read_excel(caminho, header=header_row, dtype=str)
        elif ext == ".csv":
            return pd.read_csv(caminho, header=header_row, dtype=str,
                               sep=None, engine="python")
        else:
            print(f"❌ Formato não suportado: {ext}. Use .xlsx, .xls ou .csv")
            sys.exit(1)
    except Exception as e:
        print(f"❌ Erro ao ler ficheiro: {e}")
        sys.exit(1)


def ler_coluna(caminho: str, coluna=0, header: bool = True) -> list:
    """
    Lê uma coluna específica de um ficheiro e devolve lista de valores limpos.

    Parâmetros
    ----------
    caminho : str
        Caminho absoluto do ficheiro
    coluna : int ou str
        Índice (0, 1, 2...) ou nome da coluna se header=True
    header : bool
        Se a primeira linha é cabeçalho

    Exemplo
    -------
    entidades = ler_coluna("entidades.xlsx", coluna=0, header=False)
    # → ['123', '456', '789']
    """
    df = ler_dataframe(caminho, header=header)

    try:
        col = df.iloc[:, coluna] if isinstance(coluna, int) else df[coluna]
    except (IndexError, KeyError):
        print(f"❌ Coluna '{coluna}' não encontrada no ficheiro.")
        sys.exit(1)

    valores = [str(v).strip() for v in col if str(v).strip() not in ("", "nan", "None")]

    if not valores:
        print("❌ A coluna está vazia.")
        sys.exit(1)

    print(f"✅ {len(valores)} valor(es) carregado(s) da coluna '{coluna}'.")
    return valores


def formatar_in_clause(valores: list, numerico: bool = False, coluna: str = None) -> str:
    """
    Formata uma lista de valores para usar num SQL IN (...).

    O Oracle tem um limite de 1000 expressões por IN (ORA-01795).
    Se a lista tiver mais de 1000 valores, esta função divide
    automaticamente em blocos e devolve a cláusula já pronta com OR,
    envolvida em parênteses — só funciona se 'coluna' for fornecido.

    Parâmetros
    ----------
    valores : list
        Lista de valores
    numerico : bool
        True  → sem aspas:  (123, 456, 789)
        False → com aspas:  ('ABC', 'DEF')
    coluna : str, opcional
        Nome da coluna SQL. Se fornecido E a lista tiver mais de 1000
        valores, devolve a cláusula completa com OR entre blocos:
            (COD_ENT IN (...) OR COD_ENT IN (...))
        Se não for fornecido, devolve só os valores formatados
        (uso normal, sem quebra automática — falha se > 1000 valores).

    Exemplos
    --------
    # Uso simples (≤ 1000 valores)
    sql_in = formatar_in_clause(['123', '456'], numerico=True)
    # → "123, 456"
    # WHERE COD_ENT IN ({sql_in})

    # Uso com quebra automática (> 1000 valores)
    where_clause = formatar_in_clause(lista_2997, numerico=True, coluna="COD_ENT")
    # → "(COD_ENT IN (1,2,...,1000) OR COD_ENT IN (1001,...,2000) OR COD_ENT IN (2001,...,2997))"
    # WHERE {where_clause}
    """
    def _formatar_valor(v):
        if numerico:
            try:
                return str(int(float(v)))
            except ValueError as e:
                print(f"❌ Valor não numérico encontrado: {e}")
                sys.exit(1)
        return f"'{v}'"

    valores_formatados = [_formatar_valor(v) for v in valores]

    # Uso simples — sem quebra automática (comportamento original)
    if coluna is None:
        if len(valores_formatados) > 1000:
            print(f"⚠️ Aviso: {len(valores_formatados)} valores excede o limite "
                  f"do Oracle (1000). Forneça o parâmetro 'coluna' para "
                  f"activar a quebra automática em blocos.")
        return ", ".join(valores_formatados)

    # Quebra automática em blocos de 1000
    BLOCO = 1000
    blocos = [
        valores_formatados[i:i + BLOCO]
        for i in range(0, len(valores_formatados), BLOCO)
    ]

    if len(blocos) == 1:
        return f"{coluna} IN ({', '.join(blocos[0])})"

    print(f"ℹ️ {len(valores_formatados)} valores divididos em {len(blocos)} "
          f"blocos (limite Oracle de 1000 por IN).")

    clausulas = [f"{coluna} IN ({', '.join(bloco)})" for bloco in blocos]
    return "(" + " OR ".join(clausulas) + ")"


def clausula_filtro_user(filtro_tipo: str, filtro_valor: str,
                         coluna_gestor: str = "GESTOR",
                         coluna_un: str = "UN") -> str:
    """
    Constrói a cláusula SQL de filtro de dados por utilizador
    (gestor de carteira ou gerente de UN).

    Parâmetros
    ----------
    filtro_tipo : str
        "normal", "gestor" ou "gerente" — recebido via sys.argv
    filtro_valor : str
        Código de gestor ou de UN — recebido via sys.argv
    coluna_gestor : str
        Nome da coluna de gestor na query (default "GESTOR")
    coluna_un : str
        Nome da coluna de UN na query (default "UN")

    Devolve
    -------
    str
        Cláusula SQL pronta para AND, ou "1=1" se não houver filtro
        (assim a query nunca fica com WHERE inválido).

    Uso no script
    -------------
        filtro_tipo  = sys.argv[-2]
        filtro_valor = sys.argv[-1]
        clausula = clausula_filtro_user(filtro_tipo, filtro_valor)

        query = f'''
            SELECT UN, GESTOR, CLIENTE, SALDO
            FROM TABELA
            WHERE {clausula}
        '''
    """
    if filtro_tipo == "gestor" and filtro_valor:
        return f"{coluna_gestor} = '{filtro_valor}'"
    if filtro_tipo == "gerente" and filtro_valor:
        return f"{coluna_un} = '{filtro_valor}'"
    return "1=1"


def limpar_dataframe(df, colunas_especificas: list = None) -> "pd.DataFrame":
    """
    Limpa todos os caracteres problemáticos de um DataFrame antes de
    escrever para Excel — caracteres de controlo, bytes nulos, SUB (0x1A),
    e qualquer caracter não imprimível que o openpyxl não consiga serializar.

    Aplica-se a todas as colunas do tipo object (texto) automaticamente,
    ou a colunas específicas se fornecidas.
    Colunas numéricas e de data não são afectadas.

    Parâmetros
    ----------
    df : pd.DataFrame
        DataFrame a limpar (não é modificado — devolve uma cópia)
    colunas_especificas : list, opcional
        Lista de nomes de colunas para limpar. Se None, limpa todas as colunas de texto.

    Devolve
    -------
    pd.DataFrame
        Cópia do DataFrame com texto limpo
    """
    
    # Fazer uma cópia para não modificar o original
    df_limpo = df.copy()
    
    def _limpar_texto(valor):
        """
        Limpa um único valor de texto, tratando diferentes tipos de dados
        que podem vir do Oracle.
        """
        # Se for None, NaN ou outro valor nulo, retornar como está
        if pd.isna(valor) or valor is None:
            return ''
        
        # Se for bytes, decodificar para string
        if isinstance(valor, bytes):
            try:
                valor = valor.decode('utf-8', errors='ignore')
            except:
                try:
                    valor = valor.decode('latin-1', errors='ignore')
                except:
                    valor = str(valor)
        
        # Converter para string
        if not isinstance(valor, str):
            valor = str(valor)
        
        # Se for vazio, retornar vazio
        if not valor:
            return ''
        
        # Remover caracteres de controle (0x00-0x1F) exceto tab, newline, carriage return
        valor = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', valor)
        
        # Remover especificamente o caractere SUB (0x1A)
        valor = valor.replace('\x1a', '').replace('\u001a', '').replace('\ufffd', '')
        
        # Remover caracteres especiais problemáticos
        valor = valor.replace('\x0b', ' ')  # Tab vertical
        valor = valor.replace('\x0c', ' ')  # Form feed
        
        # 🔥 NOVO: Substituir caracteres especiais que causam problemas no Excel
        substituicoes = {
            '→': '-',   # Seta para direita
            '←': '-',   # Seta para esquerda
            '↑': '-',   # Seta para cima
            '↓': '-',   # Seta para baixo
            '⇒': '=>',  # Seta dupla
            '⇐': '<=',  # Seta dupla esquerda
            '⇔': '<=>', # Seta dupla ambos
            '–': '-',   # Travessão curto
            '—': '-',   # Travessão longo
            '"': "'",   # Aspas curvas abertas
            '"': "'",   # Aspas curvas fechadas
            ''': "'",   # Aspa simples curva aberta
            ''': "'",   # Aspa simples curva fechada
            '•': '*',   # Bala
            '·': '.',   # Ponto médio
        }
        for antigo, novo in substituicoes.items():
            valor = valor.replace(antigo, novo)
        
        # Tentar normalizar caracteres Unicode
        try:
            valor = unicodedata.normalize('NFKC', valor)
        except:
            pass
        
        # Remover caracteres que não são imprimíveis
        # Mantém: letras, números, espaços, pontuação comum, acentos
        valor = ''.join(c for c in valor if c.isprintable() or c in '\n\r\t')
        
        # Limpar espaços extras
        valor = ' '.join(valor.split())
        
        return valor
    
    # Determinar quais colunas processar
    if colunas_especificas is not None:
        colunas_para_processar = [col for col in colunas_especificas if col in df_limpo.columns]
        if not colunas_para_processar:
            print(f"⚠️ Nenhuma das colunas especificadas {colunas_especificas} foi encontrada.")
            return df_limpo
    else:
        # Processar apenas colunas do tipo object (texto)
        colunas_para_processar = [col for col in df_limpo.columns if df_limpo[col].dtype == object]
    
    # Aplicar limpeza
    for col in colunas_para_processar:
        try:
            df_limpo[col] = df_limpo[col].apply(_limpar_texto)
        except Exception as e:
            print(f"⚠️ Erro ao limpar coluna '{col}': {e}")
            # Em caso de erro, tenta converter para string primeiro
            try:
                df_limpo[col] = df_limpo[col].astype(str).apply(_limpar_texto)
            except Exception as e2:
                print(f"❌ Não foi possível limpar a coluna '{col}'. Mantendo original. Erro: {e2}")
    
    return df_limpo


def tratar_dados_oracle(df) -> pd.DataFrame:
    """
    Função específica para tratar dados vindos do Oracle.
    Converte bytes para strings e limpa caracteres problemáticos.
    
    Parâmetros
    ----------
    df : pd.DataFrame
        DataFrame com dados do Oracle
        
    Devolve
    -------
    pd.DataFrame
        DataFrame tratado
    """
    # Fazer uma cópia para não modificar o original
    df_tratado = df.copy()
    
    # Primeiro, converter todos os tipos object para string
    for col in df_tratado.columns:
        if df_tratado[col].dtype == object:
            # Converter bytes para string
            df_tratado[col] = df_tratado[col].apply(
                lambda x: x.decode('utf-8', errors='ignore') if isinstance(x, bytes) else x
            )
            # Converter para string
            df_tratado[col] = df_tratado[col].astype(str)
    
    # Depois aplicar a limpeza completa
    return limpar_dataframe(df_tratado)