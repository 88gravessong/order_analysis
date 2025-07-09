@echo off
chcp 65001 >nul
title 快速打包订单分析系统

echo ================================
echo    一键打包订单分析系统
echo ================================
echo.

:: 激活虚拟环境并安装依赖
if not exist ".venv" (
    echo [步骤1] 创建虚拟环境...
    python -m venv .venv
)

echo [步骤2] 激活虚拟环境...
call .venv\Scripts\activate.bat

echo [步骤3] 安装/更新依赖...
pip install -r requirements.txt
pip install pyinstaller

:: 清理旧文件
echo [步骤4] 清理旧文件...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del "*.spec"
if exist "打包结果" rmdir /s /q "打包结果"

:: 开始打包
echo [步骤5] 开始打包...
echo 正在生成独立exe文件，请稍候...

pyinstaller ^
    --onefile ^
    --noconsole ^
    --name="订单分析系统" ^
    --add-data="templates;templates" ^
    --add-data="compute_logic.py;." ^
    --hidden-import=openpyxl.utils ^
    --hidden-import=openpyxl.workbook ^
    --hidden-import=openpyxl.worksheet ^
    --hidden-import=flask ^
    --hidden-import=werkzeug ^
    --distpath="./打包结果" ^
    app_standalone.py

if errorlevel 1 (
    echo.
    echo ❌ 打包失败！
    echo 请检查错误信息并重试
    pause
    exit /b 1
)

:: 显示结果
echo.
echo ================================
echo ✅ 打包成功！
echo ================================
echo.
echo 📁 文件位置: ./打包结果/订单分析系统.exe
if exist "./打包结果/订单分析系统.exe" (
    for %%A in ("./打包结果/订单分析系统.exe") do (
        set /a size=%%~zA/1024/1024
        echo 📊 文件大小: !size! MB
    )
)
echo.
echo 🚀 使用说明:
echo   1. 将exe文件发送给同事
echo   2. 双击运行即可（无需安装Python）
echo   3. 程序会自动打开浏览器
echo   4. 可在任何Windows电脑上运行
echo.

:: 询问是否测试
set /p test="是否现在测试运行? (y/n): "
if /i "%test%"=="y" (
    echo 正在启动测试...
    start "" "./打包结果/订单分析系统.exe"
)

pause 