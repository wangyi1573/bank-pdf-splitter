@echo off
chcp 65001 >nul
setlocal

echo ========================================
echo   银行PDF回单拆分工具 - 打包脚本
echo ========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 python，请先安装 Python 3.10+ 并加入 PATH
    pause & exit /b 1
)

:: 依赖检查（缺失才安装）
python -c "import pdfplumber, pymupdf, tkinterdnd2" >nul 2>nul
if errorlevel 1 (
    echo [INFO] 安装运行依赖...
    python -m pip install pdfplumber PyMuPDF tkinterdnd2
)

python -c "import PyInstaller" >nul 2>nul
if errorlevel 1 (
    echo [INFO] 安装 PyInstaller...
    python -m pip install pyinstaller
)

echo.
echo [INFO] 开始打包...
python -m PyInstaller ^
  --name "银行PDF回单拆分工具" ^
  --onefile ^
  --windowed ^
  --clean ^
  --noconfirm ^
  --collect-all tkinterdnd2 ^
  main.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请查看上方报错信息
    pause & exit /b 1
)

echo.
echo ========================================
echo   打包完成
echo   输出: dist\银行PDF回单拆分工具.exe
echo ========================================
pause
