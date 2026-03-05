@echo off
chcp 65001 > nul
echo ============================================
echo  長照金額計算程式
echo ============================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先執行 install.bat 安裝環境。
    pause
    exit /b 1
)

cd /d "%~dp0"
python process.py
