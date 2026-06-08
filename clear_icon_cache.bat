@echo off
echo 彻底清除 Windows 图标缓存...
echo.

echo 步骤1: 终止资源管理器...
taskkill /f /im explorer.exe >nul 2>&1

echo 步骤2: 删除图标缓存文件...
del /f /s /q /a "%localappdata%\IconCache.db" >nul 2>&1
del /f /s /q /a "%localappdata%\Microsoft\Windows\Explorer\iconcache*" >nul 2>&1

echo 步骤3: 重启资源管理器...
start explorer.exe

echo.
echo 完成！现在桌面上的图标应该清晰了。
echo 如果还是糊，把exe从桌面删掉，再从文件夹重新复制到桌面。
pause