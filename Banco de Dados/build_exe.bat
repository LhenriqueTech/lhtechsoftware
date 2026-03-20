@echo off
setlocal EnableExtensions
pushd "%~dp0"

REM ====== 1) Garantir Python 3.12 ======
py -3.12 --version || (
  echo [ERRO] Python 3.12 nao encontrado.
  pause
  exit /b 1
)

REM ====== 2) Validar arquivos ======
for %%F in (app.py lh_processor.py jornada_dialog.py styles.py) do (
  if not exist "%%F" (
    echo [ERRO] Arquivo %%F nao encontrado.
    pause & exit /b 1
  )
)

REM ====== 3) Limpar build ======
rmdir /S /Q build 2>nul
rmdir /S /Q dist 2>nul
del *.spec 2>nul
del build_log.txt 2>nul

REM ====== 4) Dependencias ======
py -3.12 -m pip install --upgrade pip
if exist requirements.txt py -3.12 -m pip install -r requirements.txt
py -3.12 -m pip install pyinstaller

REM ====== 5) Icone (opcional) ======
set "ICON_OPT="
if exist "logo.ico" set "ICON_OPT=--icon=logo.ico"

REM ====== 6) Dados opcionais ======
set "DATA_OPT="
if exist "logo.png" set "DATA_OPT=--add-data logo.png;."

echo [BUILD] Iniciando build...
py -3.12 -m PyInstaller ^
  --onefile --windowed ^
  --name GeradorRelatoriosLHTech ^
  %ICON_OPT% ^
  %DATA_OPT% ^
  --hidden-import lh_processor ^
  --hidden-import jornada_dialog ^
  --collect-data openpyxl ^
  --paths "%cd%" ^
  --clean --noconfirm ^
  app.py

echo.
if exist "dist\GeradorRelatoriosLHTech.exe" (
  echo [OK] Executavel gerado com sucesso!
) else (
  echo [ERRO] Build falhou.
)
pause
