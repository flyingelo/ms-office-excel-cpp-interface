@ECHO OFF
SETLOCAL

call %~dp0\..\build\tests\Tests.exe"
exit /B %ERRORLEVEL%
