"""
tasks.py
--------
Tarefas Celery para execução assíncrona dos scripts de relatório.
Cada tarefa corre o script como subprocesso e guarda o resultado em Redis.
"""

import subprocess
import sys
import os
import time
import glob
from pathlib import Path
from celery_app import app


@app.task(bind=True, name="executar_relatorio")
def executar_relatorio_task(self, script_path: str, modo_data: str,
                             data_inicio: str = None, data_fim: str = None,
                             ficheiro_upload: str = None,
                             filtro_tipo: str = "normal", filtro_valor: str = None,
                             username: str = "anonimo", valor_filtro: str = None,
                             valor_filtro_sdata: str = None):
    """
    Executa um script de relatório como subprocesso.
    """
    base_dir   = Path(__file__).parent
    script_abs = base_dir / script_path

    # Subpasta por utilizador
    pasta_temp = base_dir / "temp" / username
    pasta_temp.mkdir(parents=True, exist_ok=True)

    # Limpar ficheiros antigos
    agora = time.time()
    for f in pasta_temp.glob("*.xlsx"):
        try:
            if (agora - os.path.getmtime(f)) > 30 * 60:
                f.unlink()
        except Exception:
            pass

    self.update_state(state="STARTED", meta={"log": [], "status": "A iniciar..."})

    if not script_abs.exists():
        raise FileNotFoundError(f"Script não encontrado: {script_abs}")

    # ── Montar argumentos ──
    if modo_data == "simples":
        args = [sys.executable, str(script_abs), data_inicio, str(pasta_temp),
                filtro_tipo, filtro_valor or ""]

    elif modo_data == "sem_filtros":
            args = [sys.executable, str(script_abs), str(pasta_temp),
                    filtro_tipo, filtro_valor or ""]
    
    elif modo_data == "periodo":
        args = [sys.executable, str(script_abs), data_inicio, data_fim, str(pasta_temp),
                filtro_tipo, filtro_valor or ""]
        
    elif modo_data == "periodo_upload":
        if not ficheiro_upload:
            raise ValueError("Ficheiro de upload obrigatório.")
        args = [sys.executable, str(script_abs), data_inicio, data_fim, ficheiro_upload, str(pasta_temp),
                filtro_tipo, filtro_valor or ""]

    elif modo_data == "simples_upload":
            if not ficheiro_upload:
                raise ValueError("Ficheiro de upload obrigatório.")
            args = [sys.executable, str(script_abs), data_inicio, ficheiro_upload, str(pasta_temp),
                    filtro_tipo, filtro_valor or ""]
            
    elif modo_data == "upload_ficheiro":
        if not ficheiro_upload:
            raise ValueError("Ficheiro de upload obrigatório.")
        args = [sys.executable, str(script_abs), ficheiro_upload, str(pasta_temp),
                filtro_tipo, filtro_valor or ""]
        
    elif modo_data == "periodo_numero":
        if not valor_filtro:
            raise ValueError("Este relatório requer um valor de filtro.")
        if not valor_filtro.isdigit():
            raise ValueError("O valor de filtro deve conter apenas dígitos.")
        args = [sys.executable, str(script_abs), data_inicio, data_fim, valor_filtro, str(pasta_temp)]

    elif modo_data == "char_sem_data":
        # Relatório para garantias por tipo
        if not valor_filtro_sdata:
            raise ValueError("Este relatório requer pelo menos um TIPO_GAR.")
        
        tipos_gar = [t.strip() for t in valor_filtro_sdata.split(',') if t.strip()]
        
        if not tipos_gar:
            raise ValueError("Nenhum TIPO_GAR válido fornecido.")
        
        for tipo in tipos_gar:
            if not tipo.isdigit():
                raise ValueError(f"TIPO_GAR inválido: '{tipo}'. Deve conter apenas dígitos.")
        
        # Construir argumentos: script + tipos + pasta_temp
        args = [sys.executable, str(script_abs)] + tipos_gar + [str(pasta_temp)]
        
        self.update_state(
            state="PROGRESS",
            meta={"log": [f"🔍 Executando relatório para TIPOS_GAR: {', '.join(tipos_gar)}"], 
                  "status": "A iniciar..."}
        )
    else:
        raise ValueError(f"Modo de data desconhecido: {modo_data}")

    # Executar subprocesso
    log_linhas = []
    processo = subprocess.Popen(
        args,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(base_dir),
        env={**os.environ, "PYTHONUNBUFFERED": "1", "PYTHONIOENCODING": "utf-8"}
    )

    for linha in processo.stdout:
        linha = linha.rstrip("\n")
        if linha:
            log_linhas.append(linha)
            self.update_state(
                state="PROGRESS",
                meta={"log": log_linhas, "status": "A executar..."}
            )

    processo.wait()

    if processo.returncode != 0:
        self.update_state(
            state="FAILURE",
            meta={
                "log": log_linhas,
                "status": "Erro na execução.",
                "exc_type": "RuntimeError",
                "exc_message": f"Script terminou com erro (código {processo.returncode}).",
            }
        )
        return {
            "log": log_linhas,
            "status": "Erro na execução.",
            "excel_path": None,
        }

    # ─────────────────────────────────────────────────────────────
    # Encontrar o Excel gerado — CORRIGIDO para char_sem_data
    # ─────────────────────────────────────────────────────────────
    excel_path = None
    agora = time.time()

    # Buscar QUALQUER arquivo .xlsx recente na pasta do usuário
    ficheiros = list(pasta_temp.glob("*.xlsx"))
    recentes = [str(f) for f in ficheiros if (agora - os.path.getmtime(f)) <= 30 * 60]
    
    if recentes:
        # Pegar o mais recente
        excel_path = max(recentes, key=os.path.getmtime)
        log_linhas.append(f"📁 Arquivo encontrado: {excel_path}")
    else:
        # Se não encontrou, tentar com padrões específicos para debug
        log_linhas.append("⚠️ Nenhum arquivo Excel recente encontrado.")
        log_linhas.append(f"📂 Procurando em: {pasta_temp}")
        log_linhas.append(f"📄 Arquivos encontrados: {[f.name for f in ficheiros]}")

    log_linhas.append("✅ Relatório concluído com sucesso!")

    return {
        "log": log_linhas,
        "status": "Concluído",
        "excel_path": excel_path,
    }