@echo off
date /T >>c:\support\defrag.log
time /T >>c:\support\defrag.log
echo.
echo Defrag Process Has Commenced >>c:\support\defrag.log
echo.
echo ---------------------------- >>c:\support\defrag.log
Defrag c: -F >>c:\support\defrag.log
echo ---------------------------- >>c:\support\defrag.log
echo.
echo.
date /t >>c:\support\defrag.log
time /t >>c:\support\defrag.log
echo The Defrag Process Has Completed on>>c:\support\defrag.log
hostname >>c:\support\defrag.log
echo =================================== >>c:\support\defrag.log
echo.
pause