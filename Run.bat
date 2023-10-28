@echo off
pip install -r requirements.txt

python "%CD%\web_crawler.py"

REM del "%CD%\HP_Dock.xlsx"

REM del "%CD%\HP_DT.xlsx"

REM del "%CD%\HP_NB.xlsx"

REM del "%CD%\Lenovo_docking.xlsx"

REM del "%CD%\lenovo_DT.xlsx"

REM del "%CD%\Lenovo_NB.xlsx"

REM del "%CD%\DELL_Dock.xlsx"

REM del "%CD%\DELL_DT.xlsx"

REM del "%CD%\DELL_NB.xlsx"

