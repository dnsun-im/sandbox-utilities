echo off
for /f %%i in ('dir c:\Users\dsun\code\clarity\customextensions\*.py /b/a-d/o-a/t:c') do set LAST=%%i
set month=%date:~4,2%
set year=%date:~10,4%

if "%month%" == "12" (
   set /a "year=%year% +1"
   set month=01
)

set datestr=01-%month%-%year%

echo The most recently created file is %LAST% %datestr%