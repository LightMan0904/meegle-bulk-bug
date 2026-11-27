@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   飞书项目 - 批量建单器（本地图形界面）
echo ============================================
echo.

echo 清理可能残留的旧实例（避免网页连到旧代码）...
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -like '*webapp*server.py*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" 2>nul

where python >nul 2>nul
if errorlevel 1 (
  echo [错误] 没找到 python，请先安装 Python 3.9+ ^(安装时勾选 Add to PATH^)
  pause
  exit /b 1
)

python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo 首次运行，安装依赖 openpyxl ...
  python -m pip install --quiet openpyxl
)

where node >nul 2>nul
if errorlevel 1 (
  echo [警告] 没找到 node，调用飞书项目 API 需要 Node.js 18+，请先安装。
)

python "webapp\server.py"
pause
