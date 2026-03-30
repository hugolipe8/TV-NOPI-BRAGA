@echo off
cd /d "%~dp0"
title TV Nopi Braga

:: ─────────────────────────────────────────────────────────────────────────────
:: Verificar Python
:: ─────────────────────────────────────────────────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Instala em https://www.python.org
    pause
    exit /b 1
)

:: ─────────────────────────────────────────────────────────────────────────────
:: Instalar dependencias se necessario
:: ─────────────────────────────────────────────────────────────────────────────
echo [TV NOPI] A verificar dependencias Python...
python -c "import requests, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [TV NOPI] A instalar requests e openpyxl...
    python -m pip install requests openpyxl --quiet
    if errorlevel 1 (
        echo [ERRO] Falha ao instalar dependencias.
        pause
        exit /b 1
    )
)

:: ─────────────────────────────────────────────────────────────────────────────
:: Matar instancia anterior do servidor (porta 8765), se existir
:: ─────────────────────────────────────────────────────────────────────────────
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":8765 " ^| findstr "LISTENING"') do (
    taskkill /PID %%p /F >nul 2>&1
)

:: ─────────────────────────────────────────────────────────────────────────────
:: Iniciar servidor Python em background (janela minimizada)
:: ─────────────────────────────────────────────────────────────────────────────
echo [TV NOPI] A iniciar servidor de dados...
start "TV Nopi Server" /MIN python "%~dp0server.py"

:: Aguardar que o servidor arranque e carregue os dados
echo [TV NOPI] A aguardar servidor (5 seg)...
timeout /t 5 /nobreak >nul

:: ─────────────────────────────────────────────────────────────────────────────
:: Encontrar Chrome
:: ─────────────────────────────────────────────────────────────────────────────
set CHROME=

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
    goto :open
)
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
    goto :open
)
if exist "%LocalAppData%\Google\Chrome\Application\chrome.exe" (
    set "CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe"
    goto :open
)

echo [AVISO] Chrome nao encontrado. A tentar abrir com o browser padrao...
start "" "%~dp0index.html"
goto :done

:open
:: Construir URL do ficheiro local
set "FILE_URL=file:///%~dp0index.html"
:: Substituir barras invertidas por barras normais no URL
set "FILE_URL=%FILE_URL:\=/%"

echo [TV NOPI] A abrir slideshow em modo kiosk...
start "" "%CHROME%" --kiosk --no-first-run --disable-translate --disable-extensions --disable-infobars --disable-session-crashed-bubble --disable-features=TranslateUI "%FILE_URL%"

:done
echo [TV NOPI] Slideshow iniciado. Fecha esta janela para parar o servidor.
echo.
echo    Para sair do modo kiosk: Alt+F4 ou Ctrl+W
echo    Para parar o servidor:   fecha esta janela ou pressiona Ctrl+C
echo.
pause
taskkill /FI "WINDOWTITLE eq TV Nopi Server" /F >nul 2>&1
