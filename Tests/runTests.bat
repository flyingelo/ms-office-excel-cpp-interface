@ECHO OFF
SETLOCAL

call "%~dp0\..\build\tests\CoordsTests.exe"
if %ERRORLEVEL% NEQ 0 exit /B %ERRORLEVEL%

call "%~dp0\..\build\tests\Tests.exe"
if %ERRORLEVEL% NEQ 0 exit /B %ERRORLEVEL%

call "%~dp0\..\build\tests\PerformanceTests.exe" --rows 100 --cols 100
if %ERRORLEVEL% NEQ 0 exit /B %ERRORLEVEL%

exit /B 0
