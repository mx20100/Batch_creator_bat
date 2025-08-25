@echo off
setlocal enabledelayedexpansion

REM Find the first Excel file in the folder
for %%F in (*.xlsx) do (
    set "excel=%%F"
    set "basename=%%~nF"
    goto :foundExcel
)

:foundExcel
if not defined excel (
    echo No Excel file found in this folder.
    pause
    exit /b
)

echo Processing Excel: %excel%

REM Convert Excel to meta.csv
python -m xlsx2csv "%excel%" "meta.csv"

REM Search for a folder that has STL files
set "stlfolder="
for /d %%D in (*) do (
    dir /b "%%D\*.stl" >nul 2>&1
    if not errorlevel 1 (
        set "stlfolder=%%D"
        goto :foundSTL
    )
)

:foundSTL
if not defined stlfolder (
    echo No folder with STL files found, stopping...
    del /q "meta.csv" >nul 2>&1
    pause
    exit /b
)

echo Found STL folder: %stlfolder%

REM Copy meta.csv into the STL folder
copy /Y "meta.csv" "%stlfolder%\meta.csv" >nul

REM Create a ZIP archive with same name as Excel file
powershell -command "Compress-Archive -Path '%stlfolder%\*' -DestinationPath '%basename%.zip' -Force"

REM Cleanup: remove meta.csv from STL folder and parent folder
del /q "%stlfolder%\meta.csv" >nul 2>&1
del /q "meta.csv" >nul 2>&1

echo Created archive: %basename%.zip
echo.
echo All tasks finished!
pause
