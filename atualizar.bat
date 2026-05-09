@echo off
echo.
echo  Atualizando HTMLs (output_final + docs)
echo  ========================================
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python scripts\gerar_docs_html_evidencias.py
if errorlevel 1 (
  echo.
  echo Falha ao gerar os documentos HTML.
  pause
  exit /b 1
)
echo.
pause
