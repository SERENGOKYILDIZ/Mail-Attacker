import os
import sys
import shutil
import subprocess
import glob

try:
    import config
except ImportError:
    print("Error: Could not import 'config.py'. Make sure you are running 'build.py' from the project root.")
    sys.exit(1)

APP_NAME = config.APP_NAME
VERSION = config.APP_VERSION

def check_and_install_pyinstaller():
    try:
        import PyInstaller
        print("PyInstaller is installed.")
    except ImportError:
        print("PyInstaller not found. Installing now...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller installed successfully.")
        except subprocess.CalledProcessError as e:
            print(f"Failed to install PyInstaller: {e}")
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
        raise e

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
        raise FileNotFoundError(f"Could not find the generated executable at {dist_exe_path}.")

def main():
    try:
        check_and_install_pyinstaller()
        
        cleanup_build_files()
        
        build_executable()
        
        move_to_release()
        
    except Exception as e:
        print(f"\nBuild Process Error: {e}")
    finally:
        print("\nPerforming final cleanup block...")
        cleanup_build_files()
        print("Done.")

if __name__ == "__main__":
    main()
