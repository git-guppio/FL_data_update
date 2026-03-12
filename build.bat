@echo off
REM =============================================================================
REM build.bat - Compila main.py in un singolo exe con Nuitka
REM
REM Prerequisiti:
REM   - Ambiente virtuale attivo (.venv) con Nuitka installato:
REM       pip install nuitka
REM   - MSVC installato (Visual Studio Build Tools)
REM
REM Output: dist\main.exe
REM =============================================================================

setlocal

REM Attiva il venv se non gia' attivo
if not defined VIRTUAL_ENV (
    if exist ".venv\Scripts\activate.bat" (
        call .venv\Scripts\activate.bat
    ) else (
        echo [ERRORE] Ambiente virtuale non trovato. Eseguire prima:
        echo         python -m venv .venv
        echo         .venv\Scripts\activate
        echo         pip install -r requirements.txt
        echo         pip install nuitka
        exit /b 1
    )
)

REM Verifica che Nuitka sia installato
python -m nuitka --version >nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Nuitka non trovato. Installare con: pip install nuitka
    exit /b 1
)

echo.
echo [BUILD] Avvio compilazione con Nuitka...
echo [BUILD] Output: dist\main.exe
echo [BUILD] La prima compilazione richiede 15-30 minuti.
echo.

python -m nuitka ^
    --onefile ^
    --enable-plugin=pyqt5 ^
    --windows-disable-console ^
    --windows-icon-from-ico=dist\my_icon.ico ^
    --include-package=win32com ^
    --include-package=win32api ^
    --include-package=pywintypes ^
    --include-package=sap ^
    --include-package=core ^
    --include-package=config ^
    --output-dir=dist ^
    --output-filename=FL_data_update ^
    main.py

if errorlevel 1 (
    echo.
    echo [ERRORE] Compilazione fallita. Verificare il log sopra.
    exit /b 1
)

echo.
echo [OK] Compilazione completata: dist\FL_data_update.exe
endlocal
