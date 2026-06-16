@echo off
chcp 65001 > nul
echo ===================================
echo    配箱工具 - Windows 启动脚本
echo ===================================
echo.

REM 检查 Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 检查依赖...
python -m pip install openpyxl

echo.
echo [2/3] 启动配箱工具...
python run.py

if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出，请检查错误信息
    pause
)
