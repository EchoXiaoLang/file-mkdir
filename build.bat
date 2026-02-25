@echo off
chcp 65001 >nul
echo 使用当前环境打包 exe（请先 pip install openpyxl pyinstaller）
pyinstaller --onefile --windowed --name "表格解析建目录" ^
  --hidden-import=openpyxl ^
  main.py
echo 打包完成后 exe 在 dist\ 目录下
pause
