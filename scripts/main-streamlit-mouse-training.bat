@echo off
cd /d D:\PROJECTS\research\mouse-training-analysis
call venv\Scripts\activate
streamlit run scripts\main-streamlit-mouse-training.py --server.port 8502
