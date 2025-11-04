@echo off
setlocal

REM ===============================
REM Setup logging (UTF-8 compatible)
REM ===============================
chcp 65001 >nul
set "logfile=%~dp0converter_log.txt"
echo =============================================================== >> "%logfile%"
echo [START] %date% %time% >> "%logfile%"
echo Converter started >> "%logfile%"
echo.

REM ===============================
REM Step 1: Locate Python or use portable fallback (cached)
REM ===============================
set "PYTHON_CMD="

REM --- Try system Python (verify it actually runs and log details)
set "PYTHON_CMD="
set "PY_PATH="

for %%P in (py.exe python.exe) do (
    for /f "usebackq delims=" %%X in (`where %%P 2^>nul`) do (
        set "PY_PATH=%%X"
        echo [INFO] Found possible Python: %%X>>"%logfile%"
        REM --- Check if it's a Microsoft Store alias
        echo %%X | find /I "WindowsApps" >nul
        if %errorlevel%==0 (
            echo [WARN] Ignoring Microsoft Store alias at %%X>>"%logfile%"
            set "PY_PATH="
            set "PYTHON_CMD="
        ) else (
            REM --- Run version check to confirm it's real
            for /f "usebackq tokens=* delims=" %%V in (`"%%P --version 2^>nul"`) do (
                set "vercheck=%%V"
            )
            if defined vercheck (
                set "PYTHON_CMD=%%P"
                echo [INFO] Valid Python detected at %%X>>"%logfile%"
                goto foundPython
            ) else (
                echo [WARN] %%P did not return a valid version>>"%logfile%"
            )
        )
    )
)

REM --- Check for portable copy in tools\
set "PORTABLE_PY=%~dp0tools\python\python.exe"
if exist "%PORTABLE_PY%" (
    set "PYTHON_CMD=%PORTABLE_PY%"
    echo Using existing portable Python.
    echo [INFO] Using portable Python>>"%logfile%"
    goto foundPython
)

REM --- Check for cached portable Python ZIP in ProgramData
set "CACHE_DIR=%ProgramData%\BatchConverter\PythonCache"
set "DEST=%~dp0tools\python"
set "ZIPFILE=%temp%\python_portable.zip"

if not exist "%CACHE_DIR%" mkdir "%CACHE_DIR%" >nul 2>&1

REM Detect OS architecture
set "ARCH=amd64"
if defined PROCESSOR_ARCHITEW6432 (
    set "ARCH=amd64"
) else if "%PROCESSOR_ARCHITECTURE%"=="x86" (
    set "ARCH=win32"
)

set "PY_VER=3.12.7"
if "%ARCH%"=="amd64" (
    set "CACHE_PY=%CACHE_DIR%\python-%PY_VER%-embed-amd64.zip"
    set "PY_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-amd64.zip"
) else (
    set "CACHE_PY=%CACHE_DIR%\python-%PY_VER%-embed-win32.zip"
    set "PY_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-win32.zip"
)

echo Python not found on system. Checking cache...
echo [INFO] Checking for cached Python>>"%logfile%"

if exist "%CACHE_PY%" (
    echo Found cached portable Python at "%CACHE_PY%"
    echo [INFO] Found cached portable Python>>"%logfile%"
) else (
    echo No cached version found, downloading...
    echo [INFO] Downloading portable Python>>"%logfile%"
    powershell -Command ^
        "Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%CACHE_PY%' -UseBasicParsing"
    if not exist "%CACHE_PY%" (
        echo Failed to download portable Python.
        echo [ERROR] Download failed>>"%logfile%"
        pause
        exit /b
    )
    echo Downloaded portable Python to cache.
    echo [INFO] Cached portable Python>>"%logfile%"
)

REM --- Extract from cache
echo Extracting portable Python...
powershell -Command ^
    "Expand-Archive -Path '%CACHE_PY%' -DestinationPath '%DEST%' -Force"

REM --- Enable imports and pip support in the embedded Python
set "PTH_FILE="
for %%F in ("%DEST%\python*._pth") do set "PTH_FILE=%%~fF"
if exist "%PTH_FILE%" (
    powershell -Command ^
        "(Get-Content '%PTH_FILE%') -replace '^[#]*\s*import site','import site' | Set-Content '%PTH_FILE%'"
    echo [INFO] Enabled import site in "%PTH_FILE%">>"%logfile%"
)

set "PORTABLE_PY=%DEST%\python.exe"
if not exist "%PORTABLE_PY%" (
    echo Portable Python setup failed.
    echo [ERROR] Portable Python extraction failed>>"%logfile%"
    pause
    exit /b
)

set "PYTHON_CMD=%PORTABLE_PY%"
echo Portable Python ready.
echo [INFO] Portable Python ready>>"%logfile%"

:foundPython
echo Python found: "%PYTHON_CMD%"
echo [INFO] Python found: %PYTHON_CMD% >> "%logfile%"
"%PYTHON_CMD%" --version >> "%logfile%" 2>&1
echo.

REM ===============================
REM Step 2: Ensure pip and xlsx2csv are installed
REM ===============================
echo Checking for pip and xlsx2csv...

REM --- Step 2A: Check for pip
"%PYTHON_CMD%" -m pip --version >nul 2>&1
if errorlevel 1 (
    echo pip not found — installing...
echo [INFO] Installing pip for portable Python >> "%logfile%"
set "GETPIP=%temp%\get-pip.py"

REM --- Clean up any old file
if exist "%GETPIP%" del "%GETPIP%" >nul 2>&1

REM --- Download get-pip.py silently and properly write to file
powershell -ExecutionPolicy Bypass -NoLogo -NoProfile -Command ^
    "& { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest 'https://bootstrap.pypa.io/get-pip.py' -OutFile \"$env:TEMP\get-pip.py\" }"

REM --- Verify file exists before continuing
if not exist "%GETPIP%" (
    echo [ERROR] get-pip.py download failed >> "%logfile%"
    echo Failed to download get-pip.py.
    pause
    exit /b
)

REM --- Run pip installer inside its folder to avoid cwd issues
pushd "%temp%"
"%PYTHON_CMD%" "%GETPIP%" >> "%logfile%" 2>&1
set "pipresult=%errorlevel%"
popd

if not "%pipresult%"=="0" (
    echo [ERROR] Failed to install pip for portable Python >> "%logfile%"
    echo pip installation failed.
    pause
    exit /b
)

del "%GETPIP%" >nul 2>&1
echo [INFO] pip installed successfully >> "%logfile%"
echo pip installed successfully.
)

REM --- Step 2B: Check for xlsx2csv
"%PYTHON_CMD%" -m pip show xlsx2csv >nul 2>&1
if %errorlevel% equ 0 (
    echo xlsx2csv already present.
    echo [INFO] xlsx2csv already present >> "%logfile%"
) else (
    echo Installing xlsx2csv...
    echo [INFO] Installing xlsx2csv >> "%logfile%"
    "%PYTHON_CMD%" -m pip install xlsx2csv >> "%logfile%" 2>&1
    if %errorlevel% neq 0 (
        echo Failed to install xlsx2csv.
        echo [ERROR] xlsx2csv installation failed >> "%logfile%"
        pause
        exit /b
    )
    echo Installed xlsx2csv successfully.
    echo [INFO] Installed xlsx2csv >> "%logfile%"
)
echo.

REM ===============================
REM Step 3: Find Excel file (.xlsx or .xlsm)
REM ===============================

set "excel="

REM Search for .xlsx first
for %%F in (*.xlsx) do (
    set "excel=%%~nxF"
    goto :excelFound
)

REM If none, look for .xlsm
for %%F in (*.xlsm) do (
    set "excel=%%~nxF"
    goto :excelFound
)

echo No Excel file (.xlsx or .xlsm) found in this folder.
echo [ERROR] No Excel file found >> "%logfile%"
pause
exit /b

:excelFound
for %%A in ("%excel%") do set "basename=%%~nA"
echo Found Excel file: "%excel%"
echo [INFO] Found Excel file: %excel% >> "%logfile%"
echo.

REM ===============================
REM Step 4: Convert Excel to meta.csv
REM ===============================
echo Converting "%excel%" to meta.csv...
echo [INFO] Converting "%excel%" >> "%logfile%"
"%PYTHON_CMD%" -m xlsx2csv "%excel%" "meta.csv" >> "%logfile%" 2>&1

if not exist "meta.csv" (
    echo Conversion failed — meta.csv not created.
    echo [ERROR] Conversion failed >> "%logfile%"
    pause
    exit /b
)
echo Conversion successful.
echo [INFO] Conversion complete >> "%logfile%"
echo.

REM ===============================
REM Step 5: Validate meta.csv
REM ===============================

echo Validating meta.csv...

set "pyfile=%temp%\validate_meta_%random%.py"

> "%pyfile%" echo import csv, sys, os
>> "%pyfile%" echo required = ['batch','filename','material','part_id','copies','next_step','order_id','technology']
>> "%pyfile%" echo errors = []; fixed = 0
>> "%pyfile%" echo try:
>> "%pyfile%" echo     with open('meta.csv', newline='', encoding='utf-8-sig') as f:
>> "%pyfile%" echo         reader = csv.DictReader(f)
>> "%pyfile%" echo         rows = list(reader)
>> "%pyfile%" echo         header = reader.fieldnames
>> "%pyfile%" echo     if header is None or [h.strip().lower() for h in header[:len(required)]] != required:
>> "%pyfile%" echo         print("Header mismatch."); sys.exit(2)
>> "%pyfile%" echo     for i, row in enumerate(rows, start=2):
>> "%pyfile%" echo         if any(row.values()):
>> "%pyfile%" echo             missing = [k for k in required if not row.get(k, '').strip()]
>> "%pyfile%" echo             if missing: errors.append(f'Row {i}: missing {", ".join(missing)}')
>> "%pyfile%" echo             val = row.get('copies', '').strip()
>> "%pyfile%" echo             if val == '' or val == '0':
>> "%pyfile%" echo                 row['copies'] = '1'; fixed += 1
>> "%pyfile%" echo     if fixed:
>> "%pyfile%" echo         with open('meta.csv','w',newline='',encoding='utf-8-sig') as f:
>> "%pyfile%" echo             w = csv.DictWriter(f, fieldnames=required); w.writeheader(); w.writerows(rows)
>> "%pyfile%" echo         print(f'Corrected {fixed} row(s) with copies=0 or empty.')
>> "%pyfile%" echo     if errors:
>> "%pyfile%" echo         print('Validation failed:'); [print(' ', e) for e in errors]; sys.exit(1)
>> "%pyfile%" echo     print('Validation passed.'); sys.exit(0)
>> "%pyfile%" echo except Exception as e:
>> "%pyfile%" echo     print('Validation error:', e); sys.exit(2)

"%PYTHON_CMD%" "%pyfile%"
set "exitcode=%errorlevel%"
del "%pyfile%" >nul 2>&1

if not "%exitcode%"=="0" (
    echo Validation failed — check meta.csv
    echo [ERROR] Validation failed >> "%logfile%"
    del /q "meta.csv" >nul 2>&1
    pause
    exit /b
)

echo [INFO] Validation passed >> "%logfile%"
echo.

REM ===============================
REM Step 6: Find STL folder
REM ===============================
set "stlfolder="
for /d %%D in (*) do (
    dir /b "%%D\*.stl" >nul 2>&1
    if not errorlevel 1 (
        set "stlfolder=%%D"
        goto :foundSTL
    )
)

echo No folder with STL files found.
echo [ERROR] No STL folder found >> "%logfile%"
del /q "meta.csv" >nul 2>&1
pause
exit /b

:foundSTL
echo STL folder found: %stlfolder%
echo [INFO] STL folder found: %stlfolder% >> "%logfile%"
echo.

REM ===============================
REM Step 7: Copy meta.csv and ZIP
REM ===============================
copy /Y "meta.csv" "%stlfolder%\meta.csv" >nul
powershell -Command "Compress-Archive -Path '%stlfolder%\*' -DestinationPath '%basename%.zip' -Force" >> "%logfile%" 2>&1
if not exist "%basename%.zip" (
    echo Failed to create zip archive.
    echo [ERROR] Failed to create zip >> "%logfile%"
    pause
    exit /b
)
echo Created archive: %basename%.zip
echo [INFO] Created archive: %basename%.zip >> "%logfile%"
echo.

REM ===============================
REM Step 8: Cleanup
REM ===============================
del /q "%stlfolder%\meta.csv" >nul 2>&1
del /q "meta.csv" >nul 2>&1
echo Cleanup complete.
echo [INFO] Cleanup done >> "%logfile%"
echo [END] %date% %time% >> "%logfile%"
echo =============================================================== >> "%logfile%"
echo.
echo All tasks completed successfully!
timeout /t 3 >nul

REM Delete log file on success
del /q "%logfile%" >nul 2>&1
exit /b
