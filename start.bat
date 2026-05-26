@echo off
chcp 65001 >nul
title 公文格式核稿系统

echo ========================================
echo   公文格式核稿系统
echo ========================================
echo.

cd /d "%~dp0"

echo 正在启动服务...
echo.

python -c "import fastapi" 2>nul
if %errorlevel% neq 0 (
    echo [提示] 正在安装依赖包，请稍候...
    pip install -r requirements.txt -q
    echo.
)

start "" http://localhost:8000

python app.py

pause
