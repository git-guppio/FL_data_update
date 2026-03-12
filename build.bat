@echo off
REM =============================================================================
REM build.bat - Compila main.py in un singolo exe con Nuitka
REM
REM Prerequisiti:
REM   - Ambiente virtuale attivo (.venv) con Nuitka installato:
REM       pip install nuitka
REM   - MSVC installato (Visual Studio Build Tools)
REM
REM Output: dist\FL_data_update.exe
REM =============================================================================

setlocal

echo.
echo ============================================================
echo  FL_data_update - Build Script
echo ============================================================
echo.

REM --- Step 1: Attiva venv ---
echo [1/4] Verifica ambiente virtuale...
if not defined VIRTUAL_ENV (
    if exist ".venv\Scripts\activate.bat" (
        echo       Attivazione .venv in corso...
        call .venv\Scripts\activate.bat
        echo       Venv attivato: %VIRTUAL_ENV%
    ) else (
        echo.
        echo [ERRORE] Ambiente virtuale non trovato. Eseguire prima:
        echo         python -m venv .venv
        echo         .venv\Scripts\activate
        echo         pip install -r requirements.txt
        echo         pip install nuitka
        goto :error
    )
) else (
    echo       Venv gia' attivo: %VIRTUAL_ENV%
)

REM --- Step 2: Verifica Nuitka ---
echo.
echo [2/4] Verifica installazione Nuitka...
python -m nuitka --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERRORE] Nuitka non trovato. Installare con:
    echo         pip install nuitka
    goto :error
)
for /f "tokens=*" %%v in ('python -m nuitka --version 2^>^&1') do (
    echo       Nuitka versione: %%v
    goto :nuitka_ok
)
:nuitka_ok

REM --- Step 3: Verifica file necessari ---
echo.
echo [3/4] Verifica file necessari...
if not exist "main.py" (
    echo [ERRORE] main.py non trovato nella directory corrente.
    goto :error
)
echo       main.py                                    OK

if not exist "dist\my_icon.ico" (
    echo [ERRORE] dist\my_icon.ico non trovato. Eseguire prima make_icon.py.
    goto :error
)
echo       dist\my_icon.ico                           OK

if not exist "config\GruppoResponsabilePianificazione.csv" (
    echo [ERRORE] config\GruppoResponsabilePianificazione.csv non trovato.
    goto :error
)
echo       config\GruppoResponsabilePianificazione.csv OK

REM --- Step 4: Compilazione ---
echo.
echo [4/4] Avvio compilazione Nuitka...
echo.
echo       ATTENZIONE: la prima compilazione richiede 15-30 minuti.
echo       Il terminale puo' sembrare bloccato durante le fasi di:
echo         - Download del runtime Nuitka (solo prima volta)
echo         - Analisi delle dipendenze
echo         - Compilazione C con MSVC
echo       E' normale: attendere il completamento.
echo.
echo ============================================================
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
    --include-data-files=config\GruppoResponsabilePianificazione.csv=config\GruppoResponsabilePianificazione.csv ^
    --output-dir=dist ^
    --output-filename=FL_data_update ^
    --show-progress ^
    --show-memory ^
    main.py

if errorlevel 1 (
    echo.
    echo ============================================================
    echo [ERRORE] Compilazione fallita. Verificare il log sopra.
    echo ============================================================
    goto :error
)

echo.
echo ============================================================
echo  Compilazione completata con successo!
echo  Output: dist\FL_data_update.exe
echo ============================================================
echo.
goto :end

:error
echo.
pause
exit /b 1

:end
endlocal
