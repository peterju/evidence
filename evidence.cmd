@echo off
rem 切換工作目錄至批次檔所在目錄
cd /d "%~dp0"
::::::::::::::::: 參數設置 :::::::::::::::::
::讀取設定檔參數
FOR /F "tokens=1,2 delims==" %%i IN ('findstr /V [#\[] _config.txt') DO set %%i=%%j
for /f "tokens=2-8 delims=.:/ " %%a in ("%date% %time: =0%") do set DateNtime=%%a%%b%%d%%e
::建立日誌儲存目錄
if not exist log mkdir log
SETLOCAL ENABLEDELAYEDEXPANSION
::判斷設定檔中所列檔案與目錄是否存在
set envflag=1
if not exist %csvFile% set envflag=0&&echo CSV檔案%csvFile%不存在
if not exist %sourceFolder% set envflag=0&&echo 來源目錄%sourceFolder%不存在
if not exist %targetFolder% mkdir %targetFolder%
if not exist %targetFolder% set envflag=0&&echo 匯出目錄%targetFolder%不存在
if %envflag%==0 Exit /b

::刪除相同時間(分)的紀錄檔，避免紀錄重複
del /q log\log*_%DateNtime%.txt 2>nul
::讀取原始CSV檔，產生憑證編號檔
FOR /F "tokens=1 delims=, skip=1" %%i IN ('type %csvFile%') do echo %%i>>log\logEvidNo_%DateNtime%.txt
::讀取憑證編號檔，排序後產生憑證編號不重複檔
powershell -command "type log\logEvidNo_%DateNtime%.txt | sort -unique">log\logUniq_%DateNtime%.txt

::逐筆根據憑證編號不重複檔，若來源目錄存在憑證則搬移至目的目錄，若不存在則顯示，以上皆記錄在 log 目錄
FOR /F "tokens=1 delims=" %%i IN (log\logUniq_%DateNtime%.txt) do (
    if exist %sourceFolder%\%%i* (
        move %sourceFolder%\%%i* %targetFolder%>>log\logMove_%DateNtime%.txt
    ) else (
        echo %%i 憑證不存在
        echo %%i 憑證不存在>>log\logLost_%DateNtime%.txt
    )
)

pause