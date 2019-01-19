@echo off

set src=%1
set BASE_DIR=%~dp0
%BASE_DIR:~0,2% REM 获取当前路径

cd ".\tool\src"
"..\bin\lua.exe" "main.lua" "%BASE_DIR% %BASE_DIR%xlsx %BASE_DIR%lua"
pause