@echo off

REM Force the console to use UTF-8
chcp 65001 > nul

REM Force Python to use UTF-8 for all input/output
set PYTHONIOENCODING=utf-8

REM Activate conda
REM Adjust the path to your Anaconda/Miniconda installation if different
call C:\Users\Admin\anaconda3\Scripts\activate.bat tejomaya

REM Go to project root
cd /d "C:\Users\Admin\Desktop\Vidit\PPT-Agent copy\PPT-Agent-LLM"

REM Run the Agent
python main.py

pause
