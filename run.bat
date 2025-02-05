@echo off
cd /d "%~dp0"

:: 啟動 Go 程式，並讓它在新視窗執行
start "" go run main.go

:: 等待一段時間，讓程式有機會啟動（可調整秒數）
timeout /t 2 /nobreak >nul

:: 取得 go 程式名稱
for /f "tokens=2 delims=," %%A in ('wmic process where "commandline like '%%go run%%'" get ProcessId /format:csv ^| findstr /v "Node"') do (
    set GOPID=%%A
)

:: 如果 GOPID 存在，則關閉 Process
if defined GOPID (
    echo Terminating process: %GOPID%
    taskkill /PID %GOPID% /F
)

pause
