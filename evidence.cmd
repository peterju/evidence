@echo off
rem �����u�@�ؿ��ܧ妸�ɩҦb�ؿ�
cd /d "%~dp0"
::::::::::::::::: �ѼƳ]�m :::::::::::::::::
::Ū���]�w�ɰѼ�
FOR /F "tokens=1,2 delims==" %%i IN ('findstr /V [#\[] _config.txt') DO set %%i=%%j
SETLOCAL ENABLEDELAYEDEXPANSION
::�P�_�]�w�ɤ��ҦC�ɮ׻P�ؿ��O�_�s�b
set envflag=1
if not exist %csvFile% set envflag=0&&echo CSV�ɮ�%csvFile%���s�b
if not exist %sourceFolder% set envflag=0&&echo �ӷ��ؿ�%sourceFolder%���s�b
if not exist %targetFolder% mkdir %targetFolder%
if not exist %targetFolder% set envflag=0&&echo �ץX�ؿ�%targetFolder%���s�b
if %envflag%==0 goto end

::Ū��CSV��
del /q step*.txt 2>nul
FOR /F "tokens=1 delims=, skip=1" %%i IN ('type %csvFile%') do @echo %%i>>step1.txt
powershell -command "type step1.txt | sort -unique">step2.txt

:end
