@echo off
echo 正在关闭 Asterorbit小行星轨道分类评估工具.exe...
taskkill /F /IM "Asterorbit小行星轨道分类评估工具.exe" 2>nul
timeout /t 1 /nobreak >nul

echo 正在清理旧文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo 正在打包中文版...
pyinstaller "Asterorbit小行星轨道分类评估工具.spec"

echo 打包完成！
pause