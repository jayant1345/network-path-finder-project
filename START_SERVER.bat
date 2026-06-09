@echo off
title CPAN Network Path Finder

cd /d "D:\Network Path Finder\network-path-finder-project -180526\network-path-finder-project"

echo Starting CPAN Network Path Finder...
echo.

call "venv1\Scripts\activate.bat"

echo Server starting at http://127.0.0.1:5001
echo Press Ctrl+C to stop the server.
echo.

start "" http://127.0.0.1:5001
python server.py

pause
