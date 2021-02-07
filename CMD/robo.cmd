for /F "tokens=1-2 usebackq delims=;" %%a in (robo.txt) do (robocopy %%a %%b /copy:dat /E /S /v /XO /MT:16 /r:1 /w:1 /LOG+:robo.log)
pause
