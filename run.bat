@echo off
setlocal
pushd "%~dp0"
echo Verificando Python...
set "PY_CMD=python"
%PY_CMD% --version >nul 2>&1
if errorlevel 1 (
  py -3 --version >nul 2>&1
  if not errorlevel 1 (
    set "PY_CMD=py -3"
  ) else (
    set "PY_CMD="
    for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
      if exist "%%D\python.exe" set "PY_CMD=\"%%D\python.exe\""
    )
    if not defined PY_CMD (
      for /d %%D in ("%ProgramFiles%\Python*", "%ProgramFiles(x86)%\Python*") do (
        if exist "%%D\python.exe" set "PY_CMD=\"%%D\python.exe\""
      )
    )
  )
)

if not defined PY_CMD (
  echo Python nao encontrado. Instale o Python 3.x e tente novamente.
  popd
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Criando ambiente virtual .venv...
  %PY_CMD% -m venv .venv
)

set VENV_PY="%~dp0.venv\Scripts\python.exe"
if not exist %VENV_PY% (
  echo Falha ao criar o ambiente virtual.
  popd
  pause
  exit /b 1
)

echo Instalando dependencias...
%VENV_PY% -m pip install --upgrade pip --quiet
%VENV_PY% -m pip install pandas openpyxl colorama --quiet

echo Executando o gerador de relatorio...
%VENV_PY% create_report.py

popd
pause
