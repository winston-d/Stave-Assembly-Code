@ECHO off
set "dir1=C:\ALICE Upgrade\ITSUsoftwareCMM\ModulePosition"
set "XlstoCsv=C:\ALICE Upgrade\ITSUsoftwareCMM\ModulePosition\XlstoCsv.vbs"

:Start
cls
echo 1. test loop
echo 2. Quit
set /p choice=I choose (1,2):
if %choice%==1 goto test
if %choice%==2 exit

:test
cls
echo running loop test 
FOR %%X in ("%dir1%\*.xls") DO cscript.exe //NoLogo %XlstoCsv% "%%~fX"
echo Done
