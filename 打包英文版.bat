@echo off
echo 正在关闭 Asterorbit.exe...
taskkill /F /IM Asterorbit.exe 2>nul
timeout /t 1 /nobreak >nul

echo 正在清理旧文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo 正在打包英文版...
pyinstaller Asterorbit.spec

echo 打包完成！
pause