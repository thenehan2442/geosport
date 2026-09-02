@echo off
cd /d C:\GeoSports
 
REM Adjust this if git isn't on your PATH in Task Scheduler's environment
set GIT_EXE="C:\Program Files\Git\cmd\git.exe"
 
echo. >> geosports_log.txt
echo ==================================================== >> geosports_log.txt
echo %DATE% %TIME% - Starting scrape >> geosports_log.txt
 
"C:\Users\thenehan\AppData\Local\Python\pythoncore-3.14-64\python.exe" Geosport.py >> geosports_log.txt 2>&1
 
if %ERRORLEVEL% NEQ 0 (
    echo %DATE% %TIME% - Scrape FAILED, exit code %ERRORLEVEL%. Skipping push. >> geosports_log.txt
    exit /b 1
)
 
echo %DATE% %TIME% - Scrape succeeded. Pushing data to GitHub... >> geosports_log.txt
 
%GIT_EXE% add geosports_history.xlsx
%GIT_EXE% commit -m "Auto-update league data %DATE% %TIME%" >> geosports_log.txt 2>&1
%GIT_EXE% push origin main >> geosports_log.txt 2>&1
 
if %ERRORLEVEL% NEQ 0 (
    echo %DATE% %TIME% - git push FAILED. Check log above. >> geosports_log.txt
) else (
    echo %DATE% %TIME% - Push complete. >> geosports_log.txt
)