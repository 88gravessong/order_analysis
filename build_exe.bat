@echo off
chcp 65001 >nul
title 订单分析系统打包器

echo ================================
echo    正在打包订单分析系统...
echo ================================
echo.

:: 检查虚拟环境是否存在
if not exist ".venv\Scripts\activate.bat" (
    echo [错误] 虚拟环境不存在！请先运行 start.bat 创建环境
    pause
    exit /b 1
)

:: 激活虚拟环境
echo [步骤1] 激活虚拟环境...
call .venv\Scripts\activate.bat

:: 安装打包依赖
echo [步骤2] 安装打包工具...
pip install pyinstaller
if errorlevel 1 (
    echo [错误] PyInstaller安装失败！
    pause
    exit /b 1
)

:: 清理之前的打包结果
echo [步骤3] 清理旧文件...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del "*.spec"

:: 开始打包
echo [步骤4] 开始打包应用...
echo 这可能需要几分钟时间，请耐心等待...
echo.

pyinstaller --onefile ^
    --noconsole ^
    --name="订单分析系统" ^
    --add-data="templates;templates" ^
    --add-data="compute_logic.py;." ^
    --hidden-import=openpyxl ^
    --hidden-import=flask ^
    --hidden-import=werkzeug ^
    --hidden-import=jinja2 ^
    --hidden-import=click ^
    --hidden-import=itsdangerous ^
    --hidden-import=markupsafe ^
    --distpath="./打包结果" ^
    app_standalone.py

if errorlevel 1 (
    echo [错误] 打包失败！
    pause
    exit /b 1
)

echo.
echo ================================
echo    打包完成！
echo ================================
echo.
echo 打包结果位置: ./打包结果/订单分析系统.exe
echo 文件大小: 
for %%A in ("./打包结果/订单分析系统.exe") do echo   %%~zA 字节
echo.
echo 使用说明:
echo 1. 将 "订单分析系统.exe" 发送给同事
echo 2. 同事双击即可运行（无需安装Python）
echo 3. 程序会自动打开浏览器访问 http://localhost:5000
echo.
pause 