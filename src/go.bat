@echo off
IF "%PROCESSOR_ARCHITEW6432%"=="" GOTO native
%SystemRoot%\Sysnative\cmd.exe /c %0 %*
exit

:native
cscript "%~dp0WordStylesIndexor.vbs"
pause