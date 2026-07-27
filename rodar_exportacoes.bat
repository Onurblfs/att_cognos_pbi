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

echo O que deseja atualizar?
echo.
echo   1^) Receita  = Receita DRE PowerBI V2 ^(irat950^)
echo   2^) Fisico   = Fisico Receita - FIS 900 ^(Power BI^)
echo   3^) Receita + Fisico
echo   0^) Tudo ^(as 6 exportacoes^)
echo.
set /p "ESCOPO=Opcao: "

set "FILTRO="
set "DESC="
if "%ESCOPO%"=="1" (
  set "FILTRO=--somente Receitas"
  set "DESC=Receita DRE PowerBI V2 (irat950)"
) else if "%ESCOPO%"=="2" (
  set "FILTRO=--somente Fisicos"
  set "DESC=Fisico Receita - FIS 900 (Power BI)"
) else if "%ESCOPO%"=="3" (
  set "FILTRO=--somente Receitas,Fisicos"
  set "DESC=Receita + Fisico"
) else if "%ESCOPO%"=="0" (
  set "FILTRO="
  set "DESC=Todas as exportacoes"
) else (
  echo Opcao invalida. Use 0, 1, 2 ou 3.
  pause
  exit /b 1
)

echo.
echo Copiar os arquivos para a pasta de rede \\10.29.2.2\... ?
echo   S^) Sim ^(copia para a rede^)
echo   N^) Nao ^(so download local em downloads\^)
echo.
set /p "MOVER=Opcao [S/N]: "

set "MOVER_ARG="
if /i "%MOVER%"=="N" set "MOVER_ARG=--sem-mover"
if /i "%MOVER%"=="NAO" set "MOVER_ARG=--sem-mover"
if /i "%MOVER%"=="NÃO" set "MOVER_ARG=--sem-mover"

set "ARGS=%FILTRO% %MOVER_ARG%"

echo.
echo ------------------------------------------------------------------------
echo   Atualizar: %DESC%
if defined MOVER_ARG (
  echo   Destino:   downloads\ local ^(sem copiar para a rede^)
) else (
  echo   Destino:   pasta de rede + downloads\
)
echo   Comando:   baixar_cognos.py %ARGS%
echo ------------------------------------------------------------------------
echo.
set /p "CONF=Confirmar e iniciar? [S/N]: "
if /i not "%CONF%"=="S" if /i not "%CONF%"=="SIM" (
  echo Cancelado.
  pause
  exit /b 0
)

echo.
echo Iniciando painel de acompanhamento em outra janela...
start "Painel Cognos PBI" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0painel.ps1"

REM Pequena pausa para o painel abrir antes do Python gravar o JSON
timeout /t 2 /nobreak >nul

echo.
echo Executando...
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
