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
set "ICON_PNG=icon\3d-batch-copy-icon.png"
set "ICON_ICO=icon\3d-batch-copy-icon.ico"

if not exist "%SOURCE_FILE%" (
    echo [错误] 找不到源文件: %SOURCE_FILE%
    pause
    exit /b 1
)

echo [1/7] 检查 PyInstaller...
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

echo [2/7] 清理旧的构建文件...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "*.spec" del /q "*.spec"
if exist "%OUTPUT_NAME%" rmdir /s /q "%OUTPUT_NAME%"

echo [3/7] 准备图标文件...
set "ICON_PARAM="
if exist "%ICON_ICO%" (
    set "ICON_PARAM=--icon=%ICON_ICO%"
    echo [信息] 已找到图标文件，直接使用: %ICON_ICO%
) else if exist "%ICON_PNG%" (
    powershell -Command "Add-Type -AssemblyName System.Drawing; $png = [System.Drawing.Bitmap]::FromFile('%CD%\%ICON_PNG%'); $ico = [System.Drawing.Icon]::FromHandle($png.GetHicon()); $stream = [System.IO.File]::Create('%CD%\%ICON_ICO%'); $ico.Save($stream); $stream.Close(); $ico.Dispose(); $png.Dispose()"
    if exist "%ICON_ICO%" (
        set "ICON_PARAM=--icon=%ICON_ICO%"
        echo [信息] 已从PNG转换生成图标文件: %ICON_ICO%
    ) else (
        echo [警告] 图标转换失败，将使用默认图标
    )
) else (
    echo [警告] 未找到图标文件: %ICON_PNG%，将使用默认图标
)

echo [4/7] 开始打包 (文件夹模式, 无控制台窗口)...
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
    !ICON_PARAM! ^
    "%SOURCE_FILE%"

if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo [5/7] 复制配置文件到输出目录...
set "TARGET_DIR=%DIST_DIR%\%OUTPUT_NAME%"
if exist "..\config.ini" (
    copy /y "..\config.ini" "%TARGET_DIR%\" >nul
    echo [信息] 已复制 config.ini
)
if exist "..\Original file list.txt" (
    copy /y "..\Original file list.txt" "%TARGET_DIR%\" >nul
    echo [信息] 已复制 Original file list.txt
)

echo [6/7] 清理过程文件并移动输出目录...
move "%TARGET_DIR%" "%CD%\" >nul
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
if exist "*.spec" del /q "*.spec"

echo [7/7] 生成ZIP压缩包...
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
