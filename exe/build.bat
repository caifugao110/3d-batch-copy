@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo ============================================
echo   3D文件批量复制工具 - PyInstaller 打包脚本
echo   模式: 文件夹模式 (onedir)
echo ============================================
echo.

cd /d "%~dp0"

set "SOURCE_FILE=..\3D文件批量复制工具.py"
set "OUTPUT_NAME=3D文件批量复制工具"
set "DIST_DIR=dist"
set "BUILD_DIR=build"

if not exist "%SOURCE_FILE%" (
    echo [错误] 找不到源文件: %SOURCE_FILE%
    pause
    exit /b 1
)

echo [1/6] 检查 PyInstaller...
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [信息] PyInstaller 未安装，正在安装...
    pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller 安装失败
        pause
        exit /b 1
    )
)

echo [2/6] 清理旧的构建文件...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "*.spec" del /q "*.spec"
if exist "%OUTPUT_NAME%" rmdir /s /q "%OUTPUT_NAME%"

echo [3/6] 开始打包 (文件夹模式, 无控制台窗口)...
pyinstaller --onedir ^
    --noconsole ^
    --name "%OUTPUT_NAME%" ^
    --distpath "%DIST_DIR%" ^
    --workpath "%BUILD_DIR%" ^
    --clean ^
    --noconfirm ^
    --hidden-import=customtkinter ^
    --hidden-import=requests ^
    --collect-all=customtkinter ^
    "%SOURCE_FILE%"

if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo [4/6] 复制配置文件到输出目录...
set "TARGET_DIR=%DIST_DIR%\%OUTPUT_NAME%"
if exist "..\config.ini" (
    copy /y "..\config.ini" "%TARGET_DIR%\" >nul
    echo [信息] 已复制 config.ini
)
if exist "..\Original file list.txt" (
    copy /y "..\Original file list.txt" "%TARGET_DIR%\" >nul
    echo [信息] 已复制 Original file list.txt
)

echo [5/6] 清理过程文件并移动输出目录...
move "%TARGET_DIR%" "%CD%\" >nul
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "*.spec" del /q "*.spec"

echo [6/6] 生成ZIP压缩包...
powershell -Command "Compress-Archive -Path '%OUTPUT_NAME%\*' -DestinationPath '%OUTPUT_NAME%.zip' -Force"
if exist "%OUTPUT_NAME%.zip" (
    echo [信息] 已生成压缩包: %OUTPUT_NAME%.zip
) else (
    echo [警告] 压缩包生成失败
)

echo.
echo ============================================
echo   打包完成!
echo   输出目录: %CD%\%OUTPUT_NAME%
echo   压缩包: %CD%\%OUTPUT_NAME%.zip
echo ============================================
echo.
