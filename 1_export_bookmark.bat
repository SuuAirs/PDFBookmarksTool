@echo off
echo ddddd 4 request
::TODO: 路径含有英文括号
chcp 65001
cd /d "%~dp0"
if '%1'=='' echo Please specify input file path or folder path & pause & exit
::echo %1

::add, remove, export, format, format_add 
set MODE=export

set Bookmark_ext=txt

set EXT=pdf
if "%MODE%"=="format" set EXT=txt

dir /ad %1 >nul 2>nul && goto DIR_ || call :Process %1 & goto End


:DIR_
set DIR=%~1
echo DIR="%DIR%"
::for /r "%DIR%" %%f in (*.*) do ( 
for %%f in ("%DIR%\*.%EXT%") do ( 
    rem echo %%f
    call :Process "%%f"
)
goto End


:Process
python bookmark_tool.py -mode=export -i="%~dpn1.pdf" -o="%~dpn1.%Bookmark_ext%" -y
if not %errorlevel%==0 set exist_error=1
goto:eof


:end
if not defined exist_error (
  echo 操作成功，5 s后自动退出
  choice /t 5 /d y /n >nul
  echo ..
) else (
  echo 存在一些错误，请回看控制台记录
  pause)

