@echo off
chcp 65001 >nul
echo ============================================================
echo Excel/CSV文件合并工具 - 打包脚本
echo ============================================================
echo.

REM 检查 Python 是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到 Python，请先安装 Python
    pause
    exit /b 1
)

echo 检查并安装依赖...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

echo.
echo 开始打包...
python build.py

echo.
echo 打包完成！
pause


