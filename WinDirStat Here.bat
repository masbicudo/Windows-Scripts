:::::::::::::::::::::::::::::::::::::::::
:: Automatically check & get admin rights
:::::::::::::::::::::::::::::::::::::::::
@echo off
CLS
ECHO.
ECHO =============================
ECHO Running Admin shell
ECHO =============================

:checkPrivileges
NET FILE 1>NUL 2>NUL
if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges )

:getPrivileges
if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges)
ECHO.
ECHO **************************************
ECHO Invoking UAC for Privilege Escalation
ECHO **************************************

setlocal DisableDelayedExpansion
set "batchPath=%~s0"
setlocal EnableDelayedExpansion
ECHO Set UAC = CreateObject^("Shell.Application"^) > "%temp%\OEgetPrivileges.vbs"
ECHO args = "ELEV " >> "%temp%\OEgetPrivileges.vbs"
ECHO For Each strArg in WScript.Arguments >> "%temp%\OEgetPrivileges.vbs"
ECHO args = args ^& strArg ^& " "  >> "%temp%\OEgetPrivileges.vbs"
ECHO Next >> "%temp%\OEgetPrivileges.vbs"
ECHO UAC.ShellExecute "!batchPath!", args, "", "runas", 1 >> "%temp%\OEgetPrivileges.vbs"
"%SystemRoot%\System32\WScript.exe" "%temp%\OEgetPrivileges.vbs" %*
exit /B

:gotPrivileges
if '%1'=='ELEV' shift /1
setlocal & pushd .
cd /d %~dp0

::::::::::::::::::::::::::::
::START
::::::::::::::::::::::::::::
@ECHO OFF
CLS
rem Searching "windirstat.exe" in common locations
if exist "%ProgramFiles%\WinDirStat\windirstat.exe" (
    call :reg "%ProgramFiles%\WinDirStat\windirstat.exe"

) else if exist "%ProgramFiles(x86)%\WinDirStat\windirstat.exe" (
    call :reg "%ProgramFiles(x86)%\WinDirStat\windirstat.exe"

) else if exist "%USERPROFILE%\PortableApps\WinDirStatPortable\App\WinDirStat\windirstat.exe" (
    call :reg "%USERPROFILE%\PortableApps\WinDirStatPortable\App\WinDirStat\windirstat.exe"

) else (
    set /p __pathName=Enter "windirstat.exe" path:%=%
    if exist "%__pathName%\windirstat.exe" (
        call :reg "%__pathName%\windirstat.exe"

    ) else if exist "%__pathName%windirstat.exe" (
        call :reg "%__pathName%windirstat.exe"

    ) else if exist "%__pathName%" (
        call :reg "%__pathName%"
    )
)
goto :eof

:reg

echo %1

set __path=%1
set __path=%__path:"=\"%

REG ADD "HKCR\Directory\shell\windirstat" /ve /d "WinDirStat Here" /f
REG ADD "HKCR\Directory\shell\windirstat\command" /ve /d "%__path% \"%%1\"" /f
REG ADD "HKCR\Directory\shell\windirstat" /v "Icon" /d "%__path%,0" /f

REG ADD "HKCR\Drive\shell\windirstat" /ve /d "WinDirStat Here" /f
REG ADD "HKCR\Drive\shell\windirstat\command" /ve /d "%__path% \"%%1\"" /f
REG ADD "HKCR\Drive\shell\windirstat" /v "Icon" /d "%__path%,0" /f

goto :eof
