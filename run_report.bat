@echo off
cd /d "C:\Users\administrador\Documents\vessel_report"

for /f "tokens=1,2 delims==" %%A in (.env) do (
    if not "%%A"=="" if not "%%A:~0,1%"=="#" set "%%A=%%B"
)

python vessel_report.py >> "C:\Users\administrador\Documents\vessel_report\log.txt" 2>&1
