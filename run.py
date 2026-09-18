import subprocess
import sys
import os

def main():
    print("=====================================================")
    print("Starting Teacher Observation Report Application...")
    print("Backend:  FastAPI on port 8000")
    print("Frontend: Vite/React (will show link below)")
    print("Press Ctrl+C to stop both servers.")
    print("=====================================================\n")
    
    # Ensure directories exist to prevent confusing errors
    base_dir = os.path.dirname(os.path.abspath(__file__))
    backend_dir = os.path.join(base_dir, "backend")
    frontend_dir = os.path.join(base_dir, "frontend")

    if not os.path.exists(backend_dir) or not os.path.exists(frontend_dir):
        print("Error: Could not find 'backend' or 'frontend' directories.")
        sys.exit(1)

    # Start Backend (FastAPI with Uvicorn)
    # Using python -m uvicorn allows using the current python environment cleanly
    backend_cmd = [sys.executable, "-m", "uvicorn", "main:app", "--reload", "--port", "8000"]
    backend_process = subprocess.Popen(
        backend_cmd,
        cwd=backend_dir
    )
    
    # Start Frontend (Vite)
    frontend_cmd = ["npm", "run", "dev"]
    frontend_process = subprocess.Popen(
        frontend_cmd,
        cwd=frontend_dir,
        # Shell is required on Windows to resolve "npm" via PATH correctly
        shell=True if sys.platform == 'win32' else False
    )

    try:
        # Wait for both processes
        backend_process.wait()
        frontend_process.wait()
    except KeyboardInterrupt:
        print("\nShutting down both services gracefully...")
        
        # Terminate based on platform
        if sys.platform == 'win32':
            # On windows, terminating shell=True might not kill the child node process
            # We use taskkill to kill the entire process tree
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(backend_process.pid)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(frontend_process.pid)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            backend_process.terminate()
            frontend_process.terminate()
        
        backend_process.wait()
        frontend_process.wait()
        print("Shutdown complete. Have a great day!")

if __name__ == "__main__":
    main()
