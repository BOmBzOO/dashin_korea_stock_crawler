@echo off
call C:\Users\bombz\AppData\Local\anaconda3\Scripts\activate.bat
call conda activate py37_32
cd /D %~dp0
python main.py
pause
