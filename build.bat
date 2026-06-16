@echo off
chcp 65001 > nul
echo ===================================
echo    配箱工具 - Windows 构建脚本
echo ===================================
echo.

REM 检查 Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+
    pause
    exit /b 1
)

echo [1/4] 安装依赖...
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller

echo.
echo [2/4] 清理旧构建...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
for /d %%i in (src\__pycache__) do rmdir /s /q "%%i"

echo.
echo [3/4] 使用 PyInstaller 编译 EXE...
python -m PyInstaller --onefile --windowed --name "配箱工具" --clean run.py

if errorlevel 1 (
    echo [错误] 编译失败，请检查错误信息
    pause
    exit /b 1
)

echo.
echo [4/4] 构建完成！
echo.
echo 输出文件: dist\配箱工具.exe
echo.
dir dist\配箱工具.exe
echo.
echo ===================================
echo 构建成功！
echo ===================================
pause
