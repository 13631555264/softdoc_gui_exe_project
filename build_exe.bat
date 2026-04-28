@echo off
chcp 65001 > nul
title 软著文档生成工具 - 打包脚本

echo ==========================================
echo 软著文档生成工具 - exe打包脚本
echo ==========================================
echo.

echo [1/5] 检查Python环境...
python --version
if errorlevel 1 (
    echo 错误: 未找到Python
    pause
    exit /b 1
)

echo.
echo [2/5] 安装核心依赖...
pip install pdfplumber python-docx Pillow requests tkinterdnd2 --quiet

echo.
echo [3/5] 查找 tkdnd 库文件...
set TKDND_PATH=
for /f "delims=" %%i in ('python -c "import tkinterdnd2, os; print(os.path.dirname(tkinterdnd2.__file__))" 2^>nul') do set TKDND_PATH=%%i
echo tkdnd路径: %TKDND_PATH%

echo.
echo [4/5] 清理旧的构建文件...
if exist build rmdir /s /q build 2>nul
if exist dist rmdir /s /q dist 2>nul
if exist *.spec del /q *.spec 2>nul

echo.
echo [5/5] 打包生成exe...
pyinstaller ^
    --onedir ^
    --name="软著文档生成工具" ^
    --windowed ^
    --clean ^
    --noconfirm ^
    --exclude-module=numpy ^
    --exclude-module=pandas ^
    --exclude-module=matplotlib ^
    --exclude-module=cv2 ^
    --exclude-module=pytesseract ^
    --add-data="src\core;core" ^
    --add-data="src\gui;gui" ^
    --add-data="%TKDND_PATH%;tkinterdnd2" ^
    --hidden-import=tkinterdnd2 ^
    src\main.py

echo.
if exist "dist\软著文档生成工具\软著文档生成工具.exe" (
    echo ==========================================
    echo 打包成功！
    echo ==========================================
    echo.
    echo 输出目录: dist\软著文档生成工具\
    echo 运行: dist\软著文档生成工具\软著文档生成工具.exe
) else (
    echo ==========================================
    echo 打包失败！
    echo ==========================================
)

pause