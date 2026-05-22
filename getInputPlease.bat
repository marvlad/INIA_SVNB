@echo off
setlocal

REM ============================================================
REM Create input.dat from CSV files in input/
REM Sorted by date/time inside filename:
REM Cuantificación_DD_MM_YYYY_HH_MM_SS.csv
REM ============================================================

set "ROOT_DIR=%~dp0"
set "INPUT_DIR=%ROOT_DIR%input"
set "OUTPUT_DAT=%ROOT_DIR%input.dat"

echo.
echo ============================================================
echo Creating input.dat sorted by date
echo ============================================================
echo Input folder:
echo   %INPUT_DIR%
echo Output file:
echo   %OUTPUT_DAT%
echo ============================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$inputDir = '%INPUT_DIR%';" ^
  "$outputDat = '%OUTPUT_DAT%';" ^
  "Get-ChildItem -Path $inputDir -Filter '*.csv' |" ^
  "Where-Object { $_.Name -match '(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})_(\d{2})' } |" ^
  "Sort-Object { " ^
  "  if ($_.Name -match '(\d{2})_(\d{2})_(\d{4})_(\d{2})_(\d{2})_(\d{2})') { " ^
  "    [datetime]::new([int]$matches[3], [int]$matches[2], [int]$matches[1], [int]$matches[4], [int]$matches[5], [int]$matches[6]) " ^
  "  } " ^
  "} |" ^
  "ForEach-Object { $_.Name } |" ^
  "Set-Content -Path $outputDat -Encoding UTF8"

echo.
echo Done.
echo Created:
echo   %OUTPUT_DAT%
echo.
echo Content:
echo ------------------------------------------------------------
type "%OUTPUT_DAT%"
echo ------------------------------------------------------------
echo.

pause
