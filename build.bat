@echo off
echo 开始打包图像格式转换工具...
pyinstaller --clean build.spec
echo 打包完成！
pause