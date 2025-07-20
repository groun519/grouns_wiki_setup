@echo off
pushd "%~dp0"
echo Current directory: %CD%

echo Checking dependencies...
python -c "import pandas, openpyxl, tabulate"
if ERRORLEVEL 1 (
    echo Installing dependencies...
    python -m pip install pandas openpyxl tabulate
    if ERRORLEVEL 1 (
        echo ERROR: Failed to install dependencies.
        pause
        popd
        exit /b 1
    )
) else (
    echo Dependencies are OK.
)

echo Running excel_to_md.py...
python excel_to_md.py
if ERRORLEVEL 1 (
    echo ERROR: excel_to_md.py failed.
    pause
    popd
    exit /b 1
) else (
    echo Conversion succeeded.
)

pause
popd
