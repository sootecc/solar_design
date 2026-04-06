@echo off
cd /d %~dp0
echo 태양광 가설계 시스템 시작 중...
python -m pip install -r requirements.txt -q
streamlit run app.py
pause
