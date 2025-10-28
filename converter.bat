@echo off
setlocal
chcp 65001 >nul

REM ===============================
REM Setup logging (UTF-8 compatible)
REM ===============================
set "logfile=%~dp0converter_log.txt"
echo ===============================================================>>"%logfile%"
echo [START] %date% %time%>>"%logfile%"
echo Converter started>>"%logfile%"
echo(

REM ===============================
REM Step 1: Locate Python or py launcher
REM ===============================
set "PYTHON_CMD="
for %%P in (py.exe python.exe) do (
    where %%P >nul 2>&1
    if not errorlevel 1 (
        set "PYTHON_CMD=%%P"
        goto foundPython
    )
)

echo Python not found on this computer.
echo [ERROR] Python not found>>"%logfile%"
echo Please install Python 3.10 or newer from:
echo   https://www.python.org/downloads/
echo Make sure to check "Add Python to PATH" during installation.
pause
exit /b

:foundPython
echo Python found: "%PYTHON_CMD%"
echo [INFO] Python found: %PYTHON_CMD%>>"%logfile%"
"%PYTHON_CMD%" --version>>"%logfile%" 2>&1
echo(

REM ===============================
REM Step 2: Ensure xlsx2csv is installed
REM ===============================
echo Checking for xlsx2csv...
"%PYTHON_CMD%" -m pip show xlsx2csv >nul 2>&1
if %errorlevel% equ 0 (
    echo xlsx2csv already present.
    echo [INFO] xlsx2csv already present>>"%logfile%"
) else (
    echo Installing xlsx2csv...
    echo [INFO] Installing xlsx2csv>>"%logfile%"
    "%PYTHON_CMD%" -m pip install --user xlsx2csv>>"%logfile%" 2>&1
    if %errorlevel% neq 0 (
        echo Failed to install xlsx2csv.
        echo [ERROR] xlsx2csv installation failed>>"%logfile%"
        pause
        exit /b
    )
    echo Installed xlsx2csv.
    echo [INFO] Installed xlsx2csv>>"%logfile%"
)
echo(

REM ===============================
REM Step 3: Find Excel file (.xlsx or .xlsm)
REM ===============================
set "excel="

for %%F in (*.xlsx *.xlsm) do (
    set "excel=%%~nxF"
    goto excelFound
)

echo No Excel file (.xlsx or .xlsm) found in this folder.
echo [ERROR] No Excel file found>>"%logfile%"
pause
exit /b

:excelFound
for %%A in ("%excel%") do set "basename=%%~nA"
echo Found Excel file: "%excel%"
echo [INFO] Found Excel file: %excel%>>"%logfile%"
echo(

REM ===============================
REM Step 4: Convert Excel to meta.csv
REM ===============================
echo Converting "%excel%" to meta.csv...
echo [INFO] Converting "%excel%">>"%logfile%"
"%PYTHON_CMD%" -m xlsx2csv "%excel%" "meta.csv">>"%logfile%" 2>&1

if not exist "meta.csv" (
    echo Conversion failed - meta.csv not created.
    echo [ERROR] Conversion failed>>"%logfile%"
    pause
    exit /b
)
echo Conversion successful.
echo [INFO] Conversion complete>>"%logfile%"
echo(

REM ===============================
REM Step 5: Validate meta.csv
REM ===============================
echo Validating meta.csv...

set "pyfile=%temp%\validate_meta_%random%.py"

> "%pyfile%" echo import csv, sys, os
>>"%pyfile%" echo required = ['batch','filename','material','part_id','copies','next_step','order_id','technology']
>>"%pyfile%" echo errors = []
>>"%pyfile%" echo fixed = 0
>>"%pyfile%" echo try:
>>"%pyfile%" echo     with open('meta.csv', newline='', encoding='utf-8-sig') as f:
>>"%pyfile%" echo         reader = csv.DictReader(f)
>>"%pyfile%" echo         rows = list(reader)
>>"%pyfile%" echo         header = reader.fieldnames
>>"%pyfile%" echo     if header is None or [h.strip().lower() for h in header[:len(required)]] != required:
>>"%pyfile%" echo         print("Header mismatch."); sys.exit(2)
>>"%pyfile%" echo     for i, row in enumerate(rows, start=2):
>>"%pyfile%" echo         if any(row.values()):
>>"%pyfile%" echo             missing = [k for k in required if not row.get(k, '').strip()]
>>"%pyfile%" echo             if missing: errors.append(f"Row {i}: missing {', '.join(missing)}")
>>"%pyfile%" echo             val = row.get('copies', '').strip()
>>"%pyfile%" echo             if val == '' or val == '0':
>>"%pyfile%" echo                 row['copies'] = '1'
>>"%pyfile%" echo                 fixed += 1
>>"%pyfile%" echo     if fixed:
>>"%pyfile%" echo         with open('meta.csv','w',newline='',encoding='utf-8-sig') as f:
>>"%pyfile%" echo             w = csv.DictWriter(f, fieldnames=required)
>>"%pyfile%" echo             w.writeheader(); w.writerows(rows)
>>"%pyfile%" echo         print(f"Corrected {fixed} row(s) with copies=0 or empty.")
>>"%pyfile%" echo     if errors:
>>"%pyfile%" echo         print("Validation failed:")
>>"%pyfile%" echo         for e in errors: print(" ", e)
>>"%pyfile%" echo         sys.exit(1)
>>"%pyfile%" echo     print("Validation passed.")
>>"%pyfile%" echo     sys.exit(0)
>>"%pyfile%" echo except Exception as e:
>>"%pyfile%" echo     print("Validation error:", e)
>>"%pyfile%" echo     sys.exit(2)

"%PYTHON_CMD%" "%pyfile%"
set "exitcode=%errorlevel%"
del "%pyfile%" >nul 2>&1

if not "%exitcode%"=="0" (
    echo Validation failed - check meta.csv
    echo [ERROR] Validation failed>>"%logfile%"
    del /q "meta.csv" >nul 2>&1
    pause
    exit /b
)
echo [INFO] Validation passed>>"%logfile%"
echo(

REM ===============================
REM Step 6: Find STL folder
REM ===============================
set "stlfolder="
for /d %%D in (*) do (
    dir /b "%%D\*.stl" >nul 2>&1
    if not errorlevel 1 (
        set "stlfolder=%%D"
        goto foundSTL
    )
)

echo No folder with STL files found.
echo [ERROR] No STL folder found>>"%logfile%"
del /q "meta.csv" >nul 2>&1
pause
exit /b

:foundSTL
echo STL folder found: %stlfolder%
echo [INFO] STL folder found: %stlfolder%>>"%logfile%"
echo(

REM ===============================
REM Step 7: Copy meta.csv and ZIP
REM ===============================
copy /Y "meta.csv" "%stlfolder%\meta.csv" >nul
where powershell >nul 2>&1
if errorlevel 1 (
    echo PowerShell not found. Cannot create zip archive.
    echo [ERROR] PowerShell missing>>"%logfile%"
    pause
    exit /b
)
powershell -command "Compress-Archive -Path '%stlfolder%\*' -DestinationPath '%basename%.zip' -Force">>"%logfile%" 2>&1
if not exist "%basename%.zip" (
    echo Failed to create zip archive.
    echo [ERROR] Failed to create zip>>"%logfile%"
    pause
    exit /b
)
echo Created archive: %basename%.zip
echo [INFO] Created archive: %basename%.zip>>"%logfile%"
echo(

REM ===============================
REM Step 8: Cleanup
REM ===============================
del /q "%stlfolder%\meta.csv" >nul 2>&1
del /q "meta.csv" >nul 2>&1
echo Cleanup complete.
echo [INFO] Cleanup done>>"%logfile%"
echo [END] %date% %time%>>"%logfile%"
echo ===============================================================>>"%logfile%"
echo(
echo All tasks completed successfully!
timeout /t 3 >nul

REM Delete log file on success
del /q "%logfile%" >nul 2>&1

exit /b
