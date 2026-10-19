@echo off
chcp 65001 > nul
title Auditor Tarifário & Peak Shaving - ANEEL A4
echo =====================================================================
echo    AUDITOR TARIFÁRIO & PEAK SHAVING - SETOR ELÉTRICO BRASILEIRO
echo    Conectado a: https://englucas1.github.io/Aneel-Tarifas_v3/
echo =====================================================================
echo.
echo Iniciando o servidor web da aplicação Streamlit...
echo Acesse no seu navegador: http://localhost:8501
echo.
cd /d "%~dp0"
.\.venv\Scripts\streamlit.exe run app.py --server.port 8501 --browser.gatherUsageStats false
pause
