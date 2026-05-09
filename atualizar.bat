@echo off
echo.
echo  Atualizando texto: docx -> output_final + docs
echo  ================================================
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python scripts\atualizar_texto.py
if errorlevel 1 (
  echo.
  echo Falha ao atualizar.
  pause
  exit /b 1
)
echo.
pause
