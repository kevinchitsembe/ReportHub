"""
job_store.py
------------
Guarda e recupera o estado dos jobs por utilizador via Redis.
Permite que o utilizador feche o browser e volte a ver o estado do seu job.

Chave Redis:  job:{username}  →  {task_id, relatorio_id, submetido_em}
Expiração:    2 horas (igual ao ficheiro temp/)
"""

import json
import time
import redis
from celery_app import app as celery_app
from celery.result import AsyncResult

# Conexão Redis (mesma instância usada pelo Celery)
_redis = redis.Redis(host="localhost", port=6379, db=2, decode_responses=True)

JOB_TTL = 1800 # 2 horas em segundos


def guardar_job(username: str, task_id: str, relatorio_id: str):
    """Associa um job ao utilizador."""
    payload = json.dumps({
        "task_id":       task_id,
        "relatorio_id":  relatorio_id,
        "submetido_em":  time.time(),
    })
    _redis.setex(f"job:{username}", JOB_TTL, payload)


def obter_job(username: str) -> dict | None:
    """Devolve o job activo do utilizador, ou None se não existir."""
    raw = _redis.get(f"job:{username}")
    if not raw:
        return None
    return json.loads(raw)


def limpar_job(username: str):
    """Remove o job do utilizador."""
    _redis.delete(f"job:{username}")


def estado_job(username: str) -> dict:
    """
    Devolve o estado actual do job do utilizador.

    Retorna:
    {
        "existe":       bool,
        "estado":       "PENDING"|"STARTED"|"PROGRESS"|"SUCCESS"|"FAILURE"|"EXPIRADO",
        "log":          [...],
        "status_msg":   str,
        "excel_path":   str | None,
        "relatorio_id": str | None,
        "submetido_em": float | None,
    }
    """
    job = obter_job(username)

    if not job:
        return {"existe": False, "estado": None, "log": [],
                "status_msg": "", "excel_path": None,
                "relatorio_id": None, "submetido_em": None}

    task    = AsyncResult(job["task_id"], app=celery_app)
    estado  = task.state
    meta    = task.info or {}

    # Tarefa ainda não foi apanhada pelo worker
    if estado == "PENDING":
        return {
            "existe":       True,
            "estado":       "PENDING",
            "log":          [],
            "status_msg":   "Na fila, aguarda worker...",
            "excel_path":   None,
            "relatorio_id": job.get("relatorio_id"),
            "submetido_em": job.get("submetido_em"),
        }

    # A executar (STARTED ou PROGRESS)
    if estado in ("STARTED", "PROGRESS"):
        return {
            "existe":       True,
            "estado":       estado,
            "log":          meta.get("log", []),
            "status_msg":   meta.get("status", "A executar..."),
            "excel_path":   None,
            "relatorio_id": job.get("relatorio_id"),
            "submetido_em": job.get("submetido_em"),
        }

    # Concluído com sucesso OU erro tratado como retorno (não raise)
    if estado == "SUCCESS":
        resultado = meta if isinstance(meta, dict) else {}
        # Verificar se foi um erro tratado como retorno
        if resultado.get("excel_path") is None and resultado.get("status") == "Erro na execução.":
            return {
                "existe":       True,
                "estado":       "FAILURE",
                "log":          resultado.get("log", []),
                "status_msg":   "Erro na execução",
                "excel_path":   None,
                "relatorio_id": job.get("relatorio_id"),
                "submetido_em": job.get("submetido_em"),
            }
        return {
            "existe":       True,
            "estado":       "SUCCESS",
            "log":          resultado.get("log", []),
            "status_msg":   "Concluído",
            "excel_path":   resultado.get("excel_path"),
            "relatorio_id": job.get("relatorio_id"),
            "submetido_em": job.get("submetido_em"),
        }

    # Falhou via raise (não deveria acontecer com a nova lógica, mas por segurança)
    if estado == "FAILURE":
        if isinstance(meta, dict):
            log   = meta.get("log", [])
            erro  = meta.get("exc_message", meta.get("status", "Erro desconhecido"))
        else:
            log  = []
            erro = str(meta)
        return {
            "existe":       True,
            "estado":       "FAILURE",
            "log":          log if log else [f"❌ {erro}"],
            "status_msg":   "Erro na execução",
            "excel_path":   None,
            "relatorio_id": job.get("relatorio_id"),
            "submetido_em": job.get("submetido_em"),
        }

    # Estado desconhecido / expirado
    return {
        "existe":       True,
        "estado":       "EXPIRADO",
        "log":          [],
        "status_msg":   "Estado desconhecido ou expirado",
        "excel_path":   None,
        "relatorio_id": job.get("relatorio_id"),
        "submetido_em": job.get("submetido_em"),
    }