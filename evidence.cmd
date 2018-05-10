@echo off
rem 切換工作目錄至批次檔所在目錄
cd /d "%~dp0"
::::::::::::::::: 參數設置 :::::::::::::::::
::讀取設定檔參數
FOR /F "tokens=1,2 delims==" %%i IN ('findstr /V [#\[] _config.txt') DO set %%i=%%j
SETLOCAL ENABLEDELAYEDEXPANSION
::判斷設定檔中所列檔案與目錄是否存在
set envflag=1
if not exist %csvFile% set envflag=0&&echo CSV檔案%csvFile%不存在
if not exist %sourceFolder% set envflag=0&&echo 來源目錄%sourceFolder%不存在
if not exist %targetFolder% mkdir %targetFolder%
if not exist %targetFolder% set envflag=0&&echo 匯出目錄%targetFolder%不存在
if %envflag%==0 goto end

::讀取CSV檔
del /q step*.txt 2>nul
FOR /F "tokens=1 delims=, skip=1" %%i IN ('type %csvFile%') do @echo %%i>>step1.txt
powershell -command "type step1.txt | sort -unique">step2.txt

:end
