@echo off
SETLOCAL EnableDelayedExpansion
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (
  set "DEL=%%a"
)
echo say the name of the colors, don't read

call :ColorText 0a "0a"
call :ColorText 0C "0C"
call :ColorText 0b "0b"
echo(
call :ColorText 19 "19"
call :ColorText 2F "2F"
call :ColorText 4e "4e"

goto :eof

:ColorText
echo off
<nul set /p ".=%DEL%" > "%~2"
findstr /v /a:%1 /R "^$" "%~2" nul
del "%~2" > nul 2>&1
goto :eof