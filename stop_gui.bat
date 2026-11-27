@echo off
chcp 65001 >nul
echo 正在关闭「批量建单器」的所有旧进程...
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -like '*webapp*server.py*' } | ForEach-Object { Write-Host ('  结束 PID ' + $_.ProcessId); Stop-Process -Id $_.ProcessId -Force }"
echo 完成。现在可以重新双击 start_gui.bat。
pause
