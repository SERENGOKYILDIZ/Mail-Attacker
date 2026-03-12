import os
import sys
import shutil
import subprocess
import glob

# Try to import config to get app name and version
try:
    import config
    APP_NAME = config.APP_NAME
    VERSION = config.APP_VERSION
except ImportError:
    print("Error: Could not import 'config.py'. Make sure you are running 'build.py' from the project root.")
    sys.exit(1)

VENV_PATH = os.path.join(os.getcwd(), ".venv")
VENV_PYTHON = os.path.join(VENV_PATH, "Scripts", "python.exe")

def ensure_venv():
    """Checks if we are running in the .venv. If not, re-runs the script with .venv's python."""
    # Check if we are already in the venv
    if sys.prefix == sys.base_prefix:
        print("Not running in virtual environment. Checking for .venv...")
        if os.path.exists(VENV_PYTHON):
            print(f"Re-starting script using {VENV_PYTHON}...")
            # Re-run the script with venv python and wait for it to finish
            result = subprocess.run([VENV_PYTHON] + sys.argv)
            sys.exit(result.returncode)
        else:
            print("Error: .venv not found. Please create a virtual environment first.")
            sys.exit(1)
    else:
        print(f"Running in virtual environment: {sys.prefix}")

def check_dependencies():
    """Checks and installs missing dependencies from requirements.txt."""
    print("Checking dependencies...")
    try:
        import customtkinter
        import PyInstaller
        import win10toast
        print("All critical dependencies are present.")
    except ImportError as e:
        missing_module = str(e).split("'")[-2]
        print(f"Missing dependency: {missing_module}. Installing requirements...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
            print("Dependencies installed successfully.")
        except subprocess.CalledProcessError as err:
            print(f"Failed to install dependencies: {err}")
            sys.exit(1)

def cleanup_build_files():
    print("Cleaning up build artifacts (build, dist, .spec)...")
    for dir_name in ['build', 'dist']:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name, ignore_errors=True)
    
    for spec_file in glob.glob('*.spec'):
        try:
            os.remove(spec_file)
        except OSError:
            pass

def build_executable():
    print(f"Building standalone executable for {APP_NAME} {VERSION}...")
    
    exe_name = APP_NAME.replace(" ", "_")
    # We use sys.executable to ensure we use the venv's python to run PyInstaller
    pyinstaller_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconsole",
        "--onefile",
        f"--icon=assets/logo.ico",
        "--add-data", "assets;assets",
        "--add-data", "config.py;.",
        "--collect-all", "customtkinter",
        "--name", exe_name,
        "app.py"
    ]
    
    try:
        subprocess.run(pyinstaller_cmd, check=True)
        print("Compilation successful.")
    except subprocess.CalledProcessError as e:
        print(f"Compilation failed with error code: {e.returncode}")
        sys.exit(e.returncode)

def move_to_release():
    print("Moving executable to the release directory...")
    
    release_dir = os.path.join('releases', VERSION)
    if not os.path.exists(release_dir):
        os.makedirs(release_dir)
        
    expected_exe_name = f"{APP_NAME.replace(' ', '_')}.exe"
    dist_exe_path = os.path.join('dist', expected_exe_name)
    
    if os.path.exists(dist_exe_path):
        final_exe_name = f"{APP_NAME.replace(' ', '_')}_{VERSION}.exe"
        final_exe_path = os.path.join(release_dir, final_exe_name)
        
        if os.path.exists(final_exe_path):
            os.remove(final_exe_path)
            
        shutil.move(dist_exe_path, final_exe_path)
        print(f"Successfully packaged release at: {final_exe_path}")
    else:
        print(f"Error: Could not find the generated executable at {dist_exe_path}.")
        sys.exit(1)

def main():
    # 1. Ensure we are in the venv
    ensure_venv()
    
    # 2. Check/Install dependencies in the venv
    check_dependencies()
    
    # 3. Clean up old builds
    cleanup_build_files()
    
    # 4. Build
    try:
        build_executable()
        # 5. Move to release
        move_to_release()
    except Exception as e:
        print(f"\nBuild Process Error: {e}")
    finally:
        print("\nPerforming final cleanup...")
        cleanup_build_files()
        print("Done.")

if __name__ == "__main__":
    main()
