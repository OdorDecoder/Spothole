@echo off
echo Starting PotholeWatch on http://localhost:8080
echo Open this URL in your browser, then close this window to stop.
cd /d "%~dp0"
python -m http.server 8080
