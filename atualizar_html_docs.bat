@echo off
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python scripts\gerar_docs_html_evidencias.py
if errorlevel 1 (
  echo.
  echo Falha ao gerar os documentos HTML.
  exit /b 1
)
echo.
echo HTMLs atualizados em output_final.
endlocal
