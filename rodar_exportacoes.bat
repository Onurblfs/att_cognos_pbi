@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion
title Cognos PBI - Exportacoes
cd /d "%~dp0"

echo.
echo ========================================================================
echo   COGNOS / Power BI — Atualizacao das bases
echo ========================================================================
echo.

REM ----- Localiza Python (Anaconda corporativo primeiro) -----
set "PYTHON="
if exist "C:\ProgramData\anaconda3\python.exe" set "PYTHON=C:\ProgramData\anaconda3\python.exe"
if not defined PYTHON if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PYTHON if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
where python >nul 2>&1 && if not defined PYTHON for /f "delims=" %%P in ('where python') do (
  set "PYTHON=%%P"
  goto :py_ok
)
:py_ok
if not defined PYTHON (
  echo [ERRO] Python nao encontrado. Instale Anaconda/Python ou ajuste este BAT.
  pause
  exit /b 1
)

echo Python: %PYTHON%
echo Pasta:  %CD%
echo.

echo Escolha o modo de execucao:
echo   1^) Todas as exportacoes ^(copia para a rede^)
echo   2^) Todas as exportacoes ^(so download, --sem-mover^)
echo   3^) Somente Custos ^(IRAT.950 Custo^) --sem-mover
echo   4^) Somente Pre-Pago --sem-mover
echo   5^) Personalizado ^(voce digita os argumentos^)
echo   0^) Sair
echo.
set /p "OPCAO=Opcao: "

set "ARGS="
if "%OPCAO%"=="0" exit /b 0
if "%OPCAO%"=="1" set "ARGS="
if "%OPCAO%"=="2" set "ARGS=--sem-mover"
if "%OPCAO%"=="3" set "ARGS=--sem-mover --somente Custo"
if "%OPCAO%"=="4" set "ARGS=--sem-mover --somente Pre-Pago"
if "%OPCAO%"=="5" (
  set /p "ARGS=Argumentos para baixar_cognos.py: "
)
if not "%OPCAO%"=="1" if not "%OPCAO%"=="2" if not "%OPCAO%"=="3" if not "%OPCAO%"=="4" if not "%OPCAO%"=="5" (
  echo Opcao invalida.
  pause
  exit /b 1
)

echo.
echo Iniciando painel de acompanhamento em outra janela...
start "Painel Cognos PBI" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0painel.ps1"

REM Pequena pausa para o painel abrir antes do Python gravar o JSON
timeout /t 2 /nobreak >nul

echo.
echo Executando: baixar_cognos.py %ARGS%
echo ------------------------------------------------------------------------
"%PYTHON%" -u baixar_cognos.py %ARGS%
set "RC=%ERRORLEVEL%"

echo.
echo ------------------------------------------------------------------------
if "%RC%"=="0" (
  echo Processo finalizado com sucesso.
) else (
  echo Processo finalizado com erros ^(codigo %RC%^).
  echo Veja a pasta downloads\ e debug\ se necessario.
)
echo.
pause
exit /b %RC%
