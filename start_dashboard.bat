@echo off
chcp 65001 >nul
rem 雙擊啟動 Streamlit Dashboard（app.py），瀏覽器會自動開啟 http://localhost:8501
cd /d "%~dp0"

rem 已經在執行的話，直接開瀏覽器，不重複啟動
netstat -ano | findstr /r /c:":8501 .*LISTENING" >nul
if not errorlevel 1 (
    echo Dashboard 已在執行，開啟瀏覽器...
    start "" http://localhost:8501
    exit /b 0
)

echo 啟動 Dashboard 中，瀏覽器會自動開啟 http://localhost:8501
echo 用完關掉這個視窗即可停止。
echo.
py -3.13 -m streamlit run app.py
if errorlevel 1 (
    echo.
    echo 啟動失敗。如果是缺少套件，請先執行：py -3.13 -m pip install -r requirements.txt
    pause
)
