"""
runner.py
---------
Executa cada script de relatório como subprocesso isolado.
Suporta modo simples, período, e período com ficheiro de upload.
"""

import subprocess
import sys
import os
from pathlib import Path


def executar_relatorio(script_path: str, modo_data: str, data_inicio: str,
                       data_fim: str = None, ficheiro_upload: str = None, valor_filtro: str = None, valor_filtro_sdata: str = None):
    """
    Executa um script de relatório como subprocesso.

    Parâmetros
        ----------
        script_path : str
            Caminho relativo ao script .py
        modo_data : str
            "simples"         → data_inicio + pasta_temp
            "periodo"         → data_inicio + data_fim + pasta_temp
            "periodo_upload"  → data_inicio + data_fim + ficheiro_upload + pasta_temp
            "periodo_numero"  → data_inicio + data_fim + valor_filtro + pasta_temp
        data_inicio : str
            Data no formato YYYYMMDD
        data_fim : str, opcional
            Data de fim YYYYMMDD
        ficheiro_upload : str, opcional
            Caminho absoluto para o ficheiro carregado pelo utilizador
        valor_filtro : str, opcional
            Valor numérico introduzido pelo utilizador (usado como prefixo LIKE)
        """
    base_dir   = Path(__file__).parent
    script_abs = base_dir / script_path
    pasta_temp = base_dir / "temp"
    pasta_temp.mkdir(exist_ok=True)

    if not script_abs.exists():
        yield f"❌ Script não encontrado: {script_abs}"
        return

    if modo_data == "simples":
        args = [sys.executable, str(script_abs), data_inicio, str(pasta_temp)]

    elif modo_data == "sem_filtros":
                args = [sys.executable, str(script_abs), str(pasta_temp)]

    elif modo_data == "periodo":
        args = [sys.executable, str(script_abs), data_inicio, data_fim, str(pasta_temp)]

    elif modo_data == "periodo_upload":
        if not ficheiro_upload:
            yield "❌ Este relatório requer um ficheiro de entidades."
            return
        args = [sys.executable, str(script_abs),
                data_inicio, data_fim, ficheiro_upload, str(pasta_temp)]

    elif modo_data == "simples_upload": 
                if not ficheiro_upload:
                    yield "❌ Este relatório requer um ficheiro."
                    return
                args = [sys.executable, str(script_abs),
                        data_inicio, ficheiro_upload, str(pasta_temp)]
        
    elif modo_data == "upload_ficheiro":
        if not ficheiro_upload:
            yield "❌ Este relatório requer um ficheiro."
            return
        args = [sys.executable, str(script_abs), ficheiro_upload, str(pasta_temp)]

    elif modo_data == "periodo_numero":
            if not valor_filtro:
                yield "❌ Este relatório requer um valor de filtro."
                return
            if not valor_filtro.isdigit():
                yield "❌ O valor de filtro deve conter apenas dígitos."
                return
            args = [sys.executable, str(script_abs),
                    data_inicio, data_fim, valor_filtro, str(pasta_temp)]

    elif modo_data == "char_sem_data":
            # Novo relatório para garantirias por tipo
            if not valor_filtro_sdata:
                yield "❌ Este relatório requer pelo menos um TIPO_GAR."
                return
            
            # O valor_filtro pode ser uma string com valores separados por vírgula
            # Exemplo: "3,8" ou "3,8,5"
            tipos_gar = [t.strip() for t in valor_filtro_sdata.split(',') if t.strip()]
            
            if not tipos_gar:
                yield "❌ Nenhum TIPO_GAR válido fornecido."
                return
            
            # Validar cada tipo
            for tipo in tipos_gar:
                if not tipo.isdigit():
                    yield f"❌ TIPO_GAR inválido: '{tipo}'. Deve conter apenas dígitos."
                    return
            
            # Construir argumentos para o script
            args = [sys.executable, str(script_abs)] + tipos_gar
            
            # Se houver pasta destino, adicionar no final
            if pasta_temp:
                args.append(str(pasta_temp))
            
            yield f"🔍 Executando relatório para TIPOS_GAR: {', '.join(tipos_gar)}"
     
    else:
        yield f"❌ Modo de data desconhecido: {modo_data}"
        return

    try:
        process = subprocess.Popen(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=str(base_dir),
            env={**os.environ, "PYTHONUNBUFFERED": "1", "PYTHONIOENCODING": "utf-8"}
        )

        for linha in process.stdout:
            yield linha.rstrip("\n")

        process.wait()

        if process.returncode == 0:
            yield "✅ Relatório concluído com sucesso!"
        else:
            yield f"❌ O script terminou com erro (código {process.returncode})."

    except Exception as e:
        yield f"❌ Erro ao executar o script: {e}"