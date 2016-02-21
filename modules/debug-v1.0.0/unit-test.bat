@echo off
:main
    call test.bat > unit-test-result.txt
    fc test-result.txt unit-test-result.txt
    if errorlevel 1 goto error
goto :fim

:error
echo failed check

:fim
del unit-test-result.txt
pause