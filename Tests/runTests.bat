@ECHO OFF
SETLOCAL

call "..\Build\tests\Tests.exe"
exit /B %ERRORLEVEL%
