@echo off
title Streamlit Service

:: Desativar pedido de email (redundante, mas mantém)
set STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

:: Ir para a pasta do projeto
cd /d C:\Users\sum001781\Desktop\projeto

:: Ativar o ambiente virtual
call venv\Scripts\activate.bat

:: Iniciar o Streamlit (agora lê config.toml da pasta .streamlit)
python -m streamlit run app.py --browser.gatherUsageStats false --server.headless true --server.address 0.0.0.0 --server.port 8502