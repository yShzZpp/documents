@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   日本球场地区工具 - 首次安装
echo ========================================
echo.

if exist ".venv\Scripts\python.exe" goto install_packages

where py >nul 2>nul
if not errorlevel 1 (
    echo 正在创建本地 Python 环境，请稍候...
    py -3 -m venv ".venv"
    goto check_venv
)

where python >nul 2>nul
if not errorlevel 1 (
    echo 正在创建本地 Python 环境，请稍候...
    python -m venv ".venv"
    goto check_venv
)

echo [失败] 没有找到 Python。
echo 请先安装 Python 3.9 或更高版本，并勾选“Add Python to PATH”。
echo 详细步骤请打开：Windows使用说明.html
goto failed

:check_venv
if not exist ".venv\Scripts\python.exe" (
    echo [失败] 无法创建本地 Python 环境。
    echo 请确认 Python 安装时包含了 pip 和 venv。
    goto failed
)

:install_packages
echo 正在从本目录的 wheels 文件夹离线安装依赖...
".venv\Scripts\python.exe" -m pip install --no-index --find-links "%CD%\wheels" -r "%CD%\requirements.txt"
if errorlevel 1 goto failed

echo.
echo [成功] 安装完成。以后直接使用“开始处理Excel.bat”即可。
if /I "%~1"=="--no-pause" exit /b 0
pause
exit /b 0

:failed
echo.
echo 安装未完成，请查看上方错误信息。
if /I "%~1"=="--no-pause" exit /b 1
pause
exit /b 1
