@echo off
REM Encerra painel que ficou rodando sem janela visivel.
REM Acontece quando a janela do painel.bat e fechada pelo X em vez de Ctrl+C:
REM o processo do Python sobrevive segurando a porta 8000, e o painel seguinte
REM sobe na 8001 com a URL antiga ainda aberta no navegador.

title Automacao Financeira - Parar painel
cd /d "%~dp0"

echo.
echo   Procurando paineis em execucao...
echo.

set ACHOU=0

for %%P in (80 8000 8001 8002 8003 8004) do call :matar %%P

if "%ACHOU%"=="0" (
    echo   Nenhum painel rodando. Pode abrir o painel.bat.
) else (
    echo.
    echo   Pronto. Agora abra o painel.bat.
)

echo.
pause
exit /b 0


:matar
REM %1 = porta. Pega o PID de quem escuta nela e encerra.
for /f "tokens=5" %%A in ('netstat -ano -p TCP ^| findstr /r /c:":%1 .*LISTENING"') do (
    echo   porta %1 ocupada pelo processo %%A - encerrando
    taskkill /pid %%A /f >nul 2>&1
    if errorlevel 1 (
        echo      [falhou] rode este arquivo como administrador
    ) else (
        echo      [ok] encerrado
        set ACHOU=1
    )
)
exit /b 0
