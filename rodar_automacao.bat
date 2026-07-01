@echo off
setlocal

title Automacao de Arquivos de Campanhas
color 0A

pushd "%~dp0"
if errorlevel 1 (
    echo [ERRO] Nao foi possivel acessar a pasta do BAT: %~dp0
    pause
    exit /b 1
)

set SCRIPT=src\automacao_importacao.py
set PYTHON=python

%PYTHON% --version >nul 2>&1
if errorlevel 1 (
    echo [ERRO] Python nao encontrado. Verifique se esta instalado e no PATH.
    popd
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    echo [ERRO] Arquivo %SCRIPT% nao encontrado.
    popd
    pause
    exit /b 1
)

echo Verificando dependencias...
%PYTHON% -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERRO] Falha ao instalar dependencias.
    popd
    pause
    exit /b 1
)

echo.
echo ============================================
echo   AUTOMACAO DE ARQUIVOS - INICIANDO
echo ============================================
echo.

%PYTHON% "%SCRIPT%"

popd
pause
