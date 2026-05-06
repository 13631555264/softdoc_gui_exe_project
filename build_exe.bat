@echo off
chcp 65001 > nul
title 精简打包（无pandas）

echo ==========================================
echo 精简打包 - 使用 openpyxl 替代 pandas
echo ==========================================

:: 清理缓存
if exist "%LOCALAPPDATA%\pyinstaller" rmdir /s /q "%LOCALAPPDATA%\pyinstaller"
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

:: 查找 tkdnd
for /f "delims=" %%i in ('python -c "import tkinterdnd2, os; print(os.path.dirname(tkinterdnd2.__file__))" 2^>nul') do set TKDND_PATH=%%i

:: 打包
pyinstaller ^
    --onedir ^
    --name="软著文档生成工具" ^
    --windowed ^
    --clean ^
    --noconfirm ^
    --add-data="src\core;core" ^
    --add-data="src\gui;gui" ^
    --add-data="%TKDND_PATH%;tkinterdnd2" ^
    --hidden-import=tkinterdnd2 ^
    --hidden-import=pdfplumber ^
    --hidden-import=fitz ^
    --hidden-import=docx ^
    --hidden-import=openpyxl ^
    --hidden-import=pypinyin ^
    --exclude-module=pandas ^
    --exclude-module=numpy ^
    --exclude-module=matplotlib ^
    --exclude-module=scipy ^
    src\main.py

echo.
if exist "dist\软著文档生成工具\软著文档生成工具.exe" (
    echo 打包成功！
    powershell -command "Get-ChildItem -Path 'dist\软著文档生成工具' -Recurse | Measure-Object -Property Length -Sum | ForEach-Object { '{0:N2} MB' -f ($_.Sum / 1MB) }"
) else (
    echo 打包失败！
)

pause