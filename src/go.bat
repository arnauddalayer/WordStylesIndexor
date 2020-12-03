@echo off

if exist "C:\Windows\SysWOW64\cscript.exe" goto x64
if exist "C:\Windows\System32\cscript.exe" goto x86
goto error

:x64
echo Environnement x64
"C:\Windows\SysWOW64\cscript.exe" WordStylesIndexor.vbs
goto end

:x86
echo Environnement x86
cscript WordStylesIndexor.vbs
goto end

goto error
echo CSCRIPT.EXE introuvable

:end
echo fin
echo.
pause