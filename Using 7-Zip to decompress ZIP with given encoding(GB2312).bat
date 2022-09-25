@echo off
D:
cd "D:\Program Files\7-Zip"
REM 看你的7Z裝哪裡，我的例子是裝在D槽
REM 這個例子是直接拖zip拉到BAT就會自行在腳本當前路徑下解壓

REM 936 GB
REM 932 JP
7z.exe x ""%1"" -mcp=936
pause
