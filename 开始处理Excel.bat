@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo 检测到尚未安装依赖，正在自动执行首次安装...
    call "%~dp0首次安装依赖.bat" --no-pause
    if errorlevel 1 goto failed
)

echo ========================================
echo   日本球场地区多语言标准化工具
echo ========================================
echo.

if "%~1"=="" (
    ".venv\Scripts\python.exe" "%~dp0build_workbook.py"
) else (
    ".venv\Scripts\python.exe" "%~dp0build_workbook.py" "%~1"
)

if errorlevel 1 goto failed
echo.
echo 可以关闭本窗口，或按任意键退出。
pause >nul
exit /b 0

:failed
echo.
echo 处理没有完成。请根据上方提示检查文件，或打开 Windows使用说明.html。
pause
exit /b 1
