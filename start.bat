@echo off
echo Starting Gefen-Ksafim servers...

start "Backend" cmd /k "cd /d C:\CLAUDE_CODE\Gafni\backend && uvicorn main:app --reload --host 0.0.0.0 --port 8001"

timeout /t 2 /nobreak >nul

start "Frontend" cmd /k "cd /d C:\CLAUDE_CODE\Gafni\frontend && npm run dev -- --port 5174"

echo.
echo Servers starting...
echo Backend:  http://localhost:8001
echo Frontend: http://localhost:5174
echo.
timeout /t 4 /nobreak >nul
start http://localhost:5174
