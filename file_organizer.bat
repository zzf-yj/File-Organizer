@echo off
chcp 65001 > nul
setlocal EnableDelayedExpansion

:: 设置标题
title 文件整理工具启动器

:: 检查便携版Python文件夹是否存在
if not exist "portable_python" (
    echo 未检测到Python环境，正在下载便携版Python...
    
    :: 创建临时目录
    mkdir temp 2>nul
    
    :: 下载便携版Python
    powershell -Command "(New-Object Net.WebClient).DownloadFile('https://github.com/winpython/winpython/releases/download/6.1.20230527/Winpython64-3.10.11.1.exe', 'temp\winpython.exe')"
    
    if not exist "temp\winpython.exe" (
        echo Python下载失败！
        pause
        exit /b 1
    )
    
    echo 正在解压Python环境...
    "temp\winpython.exe" -y -o"portable_python"
    
    :: 清理临时文件
    rmdir /s /q temp
    
    :: 安装所需库
    portable_python\python-3.10.11.amd64\python.exe -m pip install tkinter
)

:: 运行Python程序
echo 启动文件整理工具...
start "" "portable_python\python-3.10.11.amd64\pythonw.exe" file_organizer.py

exit /b 0