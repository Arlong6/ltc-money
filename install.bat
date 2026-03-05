@echo off
chcp 65001 > nul
echo ============================================
echo  安裝必要套件（只需執行一次）
echo ============================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.8 以上版本。
    echo 下載網址：https://www.python.org/downloads/
    echo 安裝時請勾選 "Add Python to PATH"
    pause
    exit /b 1
)

echo 正在安裝套件...
python -m pip install pandas openpyxl --quiet
if errorlevel 1 (
    echo [錯誤] 套件安裝失敗，請確認網路連線後重試。
    pause
    exit /b 1
)

echo.
echo 安裝完成！之後請使用 run.bat 執行程式。
pause
