@echo off

REM Chrome Removal Script

:: Remove Google Chrome from the system.
"%ProgramFiles%\Google\Chrome\Application\chrome.exe" --uninstall --force-uninstall

:: Clear the local application data.
rd /s /q "%LOCALAPPDATA%\Google\Chrome"

:: Clear the user profile data.
rd /s /q "%USERPROFILE%\AppData\Local\Google\Chrome"

:: Clear the program data.
rd /s /q "%PROGRAMDATA%\Google\Chrome"

:: Fix icon cache.
DEL /f /q "%LOCALAPPDATA%\Microsoft\Windows\Explorer\iconcache*"

:: Reset ACL permissions to default.
"icacls" "%LOCALAPPDATA%\Google\Chrome" /reset
"icacls" "%USERPROFILE%\AppData\Local\Google\Chrome" /reset
"icacls" "%PROGRAMDATA%\Google\Chrome" /reset

:: Restart Windows Explorer to refresh the icon cache.
taskkill /f /im explorer.exe
start explorer.exe

echo Chrome has been successfully removed and system cleaned up.
pause