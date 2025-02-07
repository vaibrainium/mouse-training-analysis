@echo off
cd /d D:\PROJECTS\research\mouse-training-analysis
call venv\Scripts\activate
python scripts/create-daily-plots.py
python scripts/streamlit-data-generator.py