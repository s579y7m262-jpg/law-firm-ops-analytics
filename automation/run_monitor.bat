@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0.."

rem ============================================================
rem  经营监测仪表盘 - 启动器（手动运行 / 计划任务共用）
rem  双击运行      会停留显示结果
rem  计划任务调用  加 --silent，不等待按键（可与其他参数任意顺序）
rem  强制重算      加 --force，源文件没变也重算
rem ============================================================

rem 控制台是 GBK，显式指定输出编码，避免中文乱码或 UnicodeEncodeError
set PYTHONIOENCODING=gbk:replace

rem --silent 可能出现在任意位置，逐个参数扫一遍（只判 %1 会漏）
set "SILENT="
if not "%~1"=="" (
  for %%A in (%*) do (
    if /i "%%A"=="--silent" set "SILENT=1"
  )
)

set "PY="
if defined PYTHON set "PY=%PYTHON%"

if not defined PY (
  for /f "delims=" %%P in ('where python 2^>nul') do (
    if not defined PY set "PY=%%P"
  )
)

if not defined PY (
  for /f "delims=" %%P in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do (
    if not defined PY set "PY=%%P"
  )
)

if not defined PY (
  echo [错误] 未找到 Python。请安装后重试，或设环境变量 PYTHON 指向 python.exe
  if not defined SILENT pause
  exit /b 1
)

rem 解释器必须带 pandas + openpyxl，否则跑到一半才报错
"%PY%" -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
  echo [错误] 当前解释器缺少依赖：%PY%
  echo        请执行  pip install -r requirements.txt
  if not defined SILENT pause
  exit /b 2
)

echo.
echo  计算经营监测仪表盘 ...
echo  解释器：%PY%
echo.

"%PY%" "%~dp0..\src\monitor\monitor_etl.py" %*
set "RC=%errorlevel%"

echo.
if "%RC%"=="0" (
  echo  [完成] dashboards\monitor.html 已更新
  echo  页面开着的话，2 分钟内会自动刷新数据。
) else (
  echo  [失败] 详见 build\monitor.log
)

if not defined SILENT (
  echo.
  pause
)
exit /b %RC%
