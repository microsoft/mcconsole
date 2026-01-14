@echo off
REM ─── Compute MCQ_BASE as the parent of this script’s folder ───
for %%I in ("%~dp0..") do set "MCQ_BASE=%%~fI"

REM ─── Prepend our embedded Python and its Scripts folder ───
set "PATH=%MCQ_BASE%\python;%MCQ_BASE%\python\Scripts;%PATH%"

echo.
echo McQueen Python env activated:
echo   python is "%MCQ_BASE%\python\python.exe"
echo   pip    is "%MCQ_BASE%\python\Scripts\pip.exe"
echo.

where python
where pip

REM ─── Change into the “console” folder under the base ───
cd /d "%MCQ_BASE%\console"

REM ─── Run the mcconsole script ───
python mcconsole.py

REM ─── Keep this window open under the same PATH ───
cmd.exe /k
