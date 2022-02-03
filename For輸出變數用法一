@echo off
setlocal

for /F "usebackq tokens=3*" %%A in (`reg query "HKEY_USERS\S-1-5-21-106614769-3574578271-3013450065-1001\Keyboard Layout\Preload" ^ ^| findstr "00000404"`) do (
set appdir=%%A %%B

)
echo %appdir%
