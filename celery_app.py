"""
celery_app.py
-------------
Configuração central do Celery.
Arrancar o worker com:
    celery -A celery_app worker --loglevel=info --concurrency=10 -P solo
"""

from celery import Celery

# Redis como broker (fila) e backend (resultados)
app = Celery(
    "relatorios",
    broker="redis://localhost:6379/0",
    backend="redis://localhost:6379/1",
)

app.conf.update(
    # Serialização
    task_serializer="json",
    result_serializer="json",
    accept_content=["json"],

    # Resultados expiram após 2 horas (em segundos)
    result_expires=7200,

    # Timezone
    timezone="Africa/Maputo",
    enable_utc=True,

    # Evitar que tarefas fiquem presas indefinidamente
    task_soft_time_limit=1800,   # aviso aos 30 minutos
    task_time_limit=3600,        # força encerramento ao fim de 1 hora

    # Retry automático em caso de perda de conexão ao broker
    broker_connection_retry_on_startup=True,

    # Compatibilidade com Redis 3.x (desactiva comando HELLO)
    redis_backend_health_check_interval=0,
    broker_transport_options={
        "visibility_timeout": 3600,
        "socket_connect_timeout": 5,
        "socket_timeout": 5,
    },
    result_backend_transport_options={
        "socket_connect_timeout": 5,
        "socket_timeout": 5,
    },
)