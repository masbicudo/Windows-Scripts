@echo off
@CHCP 437 >nul 2>nul
ECHO.SOURCE
ECHO.XXX            > test-source.txt
ECHO.Miguel Angelo  >> test-source.txt
ECHO.XXX            >> test-source.txt
ECHO.Santos Bicudo  >> test-source.txt
ECHO.XXX            >> test-source.txt
type test-source.txt

:: TEST #1
copy /Y test-source.txt test-result-1.txt 1>nul 2>nul
ECHO.
ECHO.
ECHO.TEST #1
cscript //nologo replace.vbs test-result-1.txt XXX ABC
type test-result-1.txt
del test-result-1.txt

:: TEST #2
copy /Y test-source.txt test-result-2.txt 1>nul 2>nul
ECHO.
ECHO.
ECHO.TEST #2
cscript //nologo replace.vbs test-result-2.txt /r /ci "x([^x]+)x" " $1 "
type test-result-2.txt
del test-result-2.txt

del test-source.txt