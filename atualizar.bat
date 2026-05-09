@echo off
echo.
echo  Atualizando tudo (output_final + docs)
echo  ========================================
echo.
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

echo [1/2] Gerando HTMLs em output_final...
python scripts\gerar_docs_html_evidencias.py
if errorlevel 1 (
  echo.
  echo Falha ao gerar HTMLs em output_final.
  pause
  exit /b 1
)

echo.
echo [2/2] Publicando em docs com layout correto...
python scripts\atualizar_texto.py
if errorlevel 1 (
  echo.
  echo Falha ao publicar em docs.
  pause
  exit /b 1
)

echo.
echo ========================================
echo  Tudo atualizado! Faca commit e push.
echo ========================================
echo.
pause
