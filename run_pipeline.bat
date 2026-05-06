@echo off
REM ─────────────────────────────────────────────────────────────────────────────
REM Pipeline: raw → output_final/painel_fomento_audiovisual.html
REM ─────────────────────────────────────────────────────────────────────────────
setlocal enabledelayedexpansion

set SCRIPTS=%~dp0scripts
set PYTHON=python

echo.
echo ==========================================
echo Fomento Audiovisual — Pipeline
echo ==========================================
echo.

echo [1/7] Gerando tabela consolidada...
cd /d "%~dp0"
%PYTHON% "%SCRIPTS%\01_gerar_tabela_consolidada.py"
if errorlevel 1 (
    echo Erro no step 1!
    exit /b 1
)

echo.
echo [2/7] Gerando datasets analíticos...
%PYTHON% "%SCRIPTS%\02_gerar_datasets.py"
if errorlevel 1 (
    echo Erro no step 2!
    exit /b 1
)

echo.
echo [3/7] Painel — Critério de Seleção...
%PYTHON% "%SCRIPTS%\03_gerar_painel_criterio_selecao.py"
if errorlevel 1 (
    echo Erro no step 3!
    exit /b 1
)

echo.
echo [4/7] Painel — Produtoras...
%PYTHON% "%SCRIPTS%\04_gerar_painel_produtoras.py"
if errorlevel 1 (
    echo Erro no step 4!
    exit /b 1
)

echo.
echo [5/7] Painel — Concentração...
%PYTHON% "%SCRIPTS%\05_gerar_painel_concentracao.py"
if errorlevel 1 (
    echo Erro no step 5!
    exit /b 1
)

echo.
echo [6/7] Painel — Comparativo de Mecanismos...
%PYTHON% "%SCRIPTS%\06_gerar_painel_comparativo.py"
if errorlevel 1 (
    echo Erro no step 6!
    exit /b 1
)

echo.
echo [7/7] Painel final unificado...
%PYTHON% "%SCRIPTS%\07_gerar_painel_final.py"
if errorlevel 1 (
    echo Erro no step 7!
    exit /b 1
)

echo.
echo ==========================================
echo Pronto! Painel disponível em:
echo   output_final/painel_fomento_audiovisual.html
echo ==========================================
echo.
pause
endlocal
