@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
echo ================================================================
echo   GOE Lean Report Generator
echo ================================================================
echo.

where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js is not installed or not on PATH.
    echo Install it from https://nodejs.org (the "LTS" button^), then
    echo close this window, open a NEW one, and double-click this file again.
    pause
    exit /b 1
)

where python >nul 2>nul
if errorlevel 1 (
    where python3 >nul 2>nul
    if errorlevel 1 (
        echo [ERROR] Python is not installed or not on PATH.
        echo Install it from https://python.org ^(tick "Add python.exe to PATH"
        echo during setup^), then close this window, open a NEW one, and
        echo double-click this file again.
        pause
        exit /b 1
    )
    set PY=python3
) else (
    set PY=python
)

if not exist "node_modules" (
    echo [1/5] First-time setup: installing Node packages...
    call npm install
    if errorlevel 1 (
        echo [ERROR] npm install failed. Check your internet connection and try again.
        pause
        exit /b 1
    )
) else (
    echo [1/5] Node packages already installed - skipping.
)

%PY% -c "import PIL" >nul 2>nul
if errorlevel 1 (
    echo [1b/5] Installing Pillow ^(image library^) for Python...
    %PY% -m pip install pillow --break-system-packages >nul 2>nul
    if errorlevel 1 (
        %PY% -m pip install pillow >nul 2>nul
    )
)

if not exist "raw_data" (
    echo.
    echo [ERROR] No "raw_data" folder found next to this file.
    echo.
    echo Create a folder named raw_data here, and inside it put:
    echo   - one sub-folder per unit ^(e.g. "1_Idea Tools"^), each containing
    echo     that unit's KPI PDF, Savings PDF, kaizen PDF^(s^), owner photo
    echo   - "Units Text Information.txt"
    echo   - the GOE-level "Cumulative Outcome Table.pdf"
    echo.
    echo Then double-click this file again.
    pause
    exit /b 1
)

echo.
echo [2/5] Converting PDFs and photos into report-ready images...
%PY% scripts\prepare_assets.py raw_data prepared_assets --dpi 220
if errorlevel 1 goto :error

echo.
echo [3/5] Reading Units Text Information.txt...
set TXT_FILE=raw_data\Units Text Information.txt
if not exist "%TXT_FILE%" (
    echo [ERROR] Could not find "%TXT_FILE%"
    goto :error
)
%PY% scripts\parse_units_txt.py "%TXT_FILE%" prepared_assets\units_parsed.json
if errorlevel 1 goto :error

echo.
echo [4/5] Building config.json...
del /q "prepared_assets\goe_cumulative-*.png" >nul 2>nul
set "CUM_PDF="
for /f "delims=" %%f in ('%PY% scripts\find_cumulative_source.py raw_data 2^>nul') do if not defined CUM_PDF set "CUM_PDF=%%f"
if defined CUM_PDF (
    echo [CHECK] Cumulative source found: !CUM_PDF!
    %PY% scripts\render_cumulative_source.py "!CUM_PDF!" prepared_assets
    if errorlevel 1 (
        echo [ERROR] Cumulative source was found, but the GOE-wide Cumulative Outcome table could not be rendered.
        goto :error
    )
    if not exist "prepared_assets\goe_cumulative-1.png" (
        echo [ERROR] Cumulative source was found, but no cumulative table image was rendered.
        goto :error
    )
    echo [CHECK] Cumulative outcome pages found:
    dir /b "prepared_assets\goe_cumulative-*.png"
) else (
    echo [ERROR] Could not find the GOE cumulative outcome table in raw_data.
    echo        Put the GOE-level cumulative outcome table source anywhere inside raw_data.
    goto :error
)
%PY% scripts\build_config.py prepared_assets\units_parsed.json prepared_assets\assets_manifest.json config_auto.json --only-uploaded --cumulative-image "prepared_assets\goe_cumulative-1.png"
if errorlevel 1 goto :error

echo.
echo ================================================================
echo   config_auto.json has been created automatically from
echo   Units Text Information.txt.
echo.
echo   Profile data, contact data, unit narrative, observations,
echo   recommendations and GOE overview are populated automatically.
echo ================================================================

echo.
echo [5/5] Generating the Word report...
if not exist "assets\cover_logos_band.png" (
    echo [WARN] assets\cover_logos_band.png is missing from the project folder.
    echo         The generator will try its bundled asset fallback.
)
if not exist "out" mkdir out
node engine\generate_report.js config_auto.json out\report.docx
if errorlevel 1 goto :error

echo.
echo ================================================================
echo   DONE. Your report is at: out\report.docx
echo   (PDF export is disabled - only the Word file is generated, so
echo   nothing ever re-opens/re-saves report.docx after this point.
echo   To get a PDF: open out\report.docx in Word yourself, check the
echo   1.3 GOE-level Outcomes table looks right, then File > Save As
echo   > PDF - that way you're saving a copy, not overwriting this one.)
echo   Open it in Word. If the Index page looks empty, press Ctrl+A
echo   then F9 to update it (Word usually does this on its own).
echo ================================================================
pause
exit /b 0

:error
echo.
echo [ERROR] Something failed above - scroll up to see the message.
pause
exit /b 1
