@echo off
chcp 65001 >nul
title 订单分析系统启动器

echo ================================
echo    订单分析系统启动中...
echo ================================
echo.

:: 检查虚拟环境是否存在
if not exist ".venv\Scripts\activate.bat" (
    echo [错误] 虚拟环境不存在！
    echo 请先运行以下命令创建虚拟环境：
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

:: 激活虚拟环境
echo [步骤1] 激活虚拟环境...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo [错误] 无法激活虚拟环境！
    pause
    exit /b 1
)
echo [完成] 虚拟环境已激活

:: 检查依赖是否安装
echo.
echo [步骤2] 检查依赖...
python -c "import flask, openpyxl" 2>nul
if errorlevel 1 (
    echo [警告] 检测到缺少依赖，正在安装...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败！
        pause
        exit /b 1
    )
)
echo [完成] 所有依赖已就绪

:: 启动Flask应用
echo.
echo [步骤3] 启动Flask服务...
echo 服务地址: http://localhost:5000
echo 按 Ctrl+C 停止服务
echo.
echo ================================
echo    服务启动完成！
echo ================================
echo.

python app.py 