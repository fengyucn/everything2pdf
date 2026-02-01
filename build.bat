@echo off
chcp 65001 >nul
echo === Everything to PDF 打包脚本 (Windows) ===
echo.

REM 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python
    pause
    exit /b 1
)

REM 检查PyInstaller
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller
)

REM 清理旧的构建
echo 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM 执行打包
echo 开始打包...
pyinstaller build_windows.spec

REM 检查结果
if exist "dist\everything2pdf.exe" (
    echo.
    echo === 打包成功! ===
    echo 可执行文件: dist\everything2pdf.exe
    dir dist\everything2pdf.exe
) else (
    echo 打包失败!
    pause
    exit /b 1
)

echo.
pause
