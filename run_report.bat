@echo off
cd /d "%~dp0"
if not exist logs mkdir logs
python weekly_report.py >> logs\manual_run.log 2>&1
