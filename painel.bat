@echo off
REM Sobe o painel da Automacao Financeira e abre o navegador.
REM Feche esta janela (ou Ctrl+C) para encerrar o painel.

title Automacao Financeira - Painel
cd /d "%~dp0"

echo.
echo   Automacao Financeira - Autorizacoes de Pagamento
echo   ================================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo   [ERRO] Python nao encontrado no PATH.
    echo   Instale o Python ou abra o terminal onde ele funciona.
    echo.
    pause
    exit /b 1
)

if not exist "config\fornecedores.yaml" (
    echo   [ERRO] Falta config\fornecedores.yaml.
    echo.
    echo   Gere o registro de contas primeiro:
    echo      python scripts\gerar_registro_fornecedores.py
    echo      python scripts\enriquecer_registro.py
    echo      copy config\fornecedores.enriquecido.yaml config\fornecedores.yaml
    echo.
    pause
    exit /b 1
)

echo   Subindo o painel. O navegador abre sozinho em alguns segundos.
echo   Para encerrar: feche esta janela ou tecle Ctrl+C.
echo.

REM O painel prefere a porta 80, que deixa o endereco sem :porta no fim. Se
REM ela estiver ocupada, cai na faixa 8000 e escreve em logs\porta.txt qual
REM foi. O navegador abre pelo nome amigavel quando ele estiver no hosts
REM (scripts\configurar_endereco.py --aplicar, como administrador) e por
REM 127.0.0.1 caso contrario.
if exist "logs\porta.txt" del /q "logs\porta.txt" >nul 2>&1
start "" /b cmd /c "for /l %%i in (1,1,30) do (if exist logs\porta.txt (set /p P=<logs\porta.txt & call :abrir & exit) else timeout /t 1 /nobreak >nul)"

python -m painel

echo.
echo   Painel encerrado.
pause
exit /b 0


:abrir
REM Monta a URL: sem ":porta" quando for a 80.
findstr /c:"organizacao.financeira.local" %SystemRoot%\System32\drivers\etc\hosts >nul 2>&1
if errorlevel 1 (set HOST=127.0.0.1) else (set HOST=organizacao.financeira.local)
if "%P%"=="80" (start http://%HOST%) else (start http://%HOST%:%P%)
exit /b 0
