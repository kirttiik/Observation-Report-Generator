@echo off
echo =====================================================
echo Starting Teacher Observation Report Application...
echo Backend:  FastAPI on port 8000
echo Frontend: Vite/React
echo Close this terminal to stop both servers.
echo =====================================================

:: Start backend in a new window to keep output separate but running concurrently
:: We use the specific python path that is known to work from the IDE
echo Starting Backend...
start "Backend (FastAPI)" cmd /k "cd backend && C:\Users\kp676\AppData\Local\Programs\Python\Python313\python.exe -m uvicorn main:app --reload --port 8000"

:: Start frontend in a new window
echo Starting Frontend...
start "Frontend (Vite)" cmd /k "cd frontend && npm run dev"

echo Done! Both apps are running in separate terminal windows.
