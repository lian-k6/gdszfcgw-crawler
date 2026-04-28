@echo off
chcp 65001 >nul
echo ==========================================
echo   添加 Windows 定时任务
echo ==========================================
echo.
echo 正在创建每天定时运行爬虫的任务...
echo.

:: 获取当前目录
set SCRIPT_DIR=%~dp0
set PYTHON_PATH=%SCRIPT_DIR%venv\Scripts\pythonw.exe
set SCHEDULER_PATH=%SCRIPT_DIR%scheduler.py

:: 检查 pythonw 是否存在，否则使用系统 pythonw
if not exist "%PYTHON_PATH%" (
    set PYTHON_PATH=pythonw
)

schtasks /Create ^
    /TN "政府采购爬虫-每日定时任务" ^
    /TR "\"%PYTHON_PATH%\" \"%SCHEDULER_PATH%\"" ^
    /SC DAILY ^
    /ST 09:00 ^
    /RL HIGHEST ^
    /F

if %ERRORLEVEL% == 0 (
    echo.
    echo 任务创建成功！
    echo 每天上午 9:00 将自动运行爬虫。
    echo.
    echo 如需修改时间，请打开
    echo   控制面板 -> 管理工具 -> 任务计划程序
    echo 找到 "政府采购爬虫-每日定时任务" 进行修改。
) else (
    echo.
    echo 任务创建失败，请尝试以管理员身份运行此脚本。
)

echo.
pause
