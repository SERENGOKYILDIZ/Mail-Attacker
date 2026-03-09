# config.py

APP_NAME = "MailFlow"
APP_VERSION = "v1.4.2"

# Window settings
WINDOW_WIDTH = 1300
WINDOW_HEIGHT = 750
RESIZABLE = False

# Theme settings
DEFAULT_THEME = "dark"
ACCENT_COLOR = "blue"

# UI Padding and styling constants
PAD_X = 20
PAD_Y = 10
LABEL_PAD_Y = (10, 2)
ENTRY_PAD_Y = (0, 10)

# System settings
import os

# Default AppData location (always exists for the pointer file)
_DEFAULT_APP_DIR = os.path.join(os.environ.get("APPDATA"), APP_NAME)
if not os.path.exists(_DEFAULT_APP_DIR):
    os.makedirs(_DEFAULT_APP_DIR)

# Pointer file that stores custom data folder path
_DATA_FOLDER_FILE = os.path.join(_DEFAULT_APP_DIR, "data_folder.txt")

def _get_data_folder():
    """Read the custom data folder from the pointer file, or use default."""
    if os.path.exists(_DATA_FOLDER_FILE):
        try:
            with open(_DATA_FOLDER_FILE, "r", encoding="utf-8") as f:
                custom_path = f.read().strip()
            if custom_path and os.path.isdir(custom_path):
                app_dir = os.path.join(custom_path, APP_NAME)
                if not os.path.exists(app_dir):
                    os.makedirs(app_dir)
                return app_dir
        except Exception:
            pass
    return _DEFAULT_APP_DIR

def set_data_folder(new_path):
    """Set a new custom data folder. Pass empty string to reset to default."""
    if new_path:
        app_dir = os.path.join(new_path, APP_NAME)
        if not os.path.exists(app_dir):
            os.makedirs(app_dir)
        with open(_DATA_FOLDER_FILE, "w", encoding="utf-8") as f:
            f.write(new_path)
    else:
        # Reset to default
        if os.path.exists(_DATA_FOLDER_FILE):
            os.remove(_DATA_FOLDER_FILE)

def get_data_folder_raw():
    """Get the raw custom path (or empty string if using default)."""
    if os.path.exists(_DATA_FOLDER_FILE):
        try:
            with open(_DATA_FOLDER_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            pass
    return ""

def reload_paths():
    """Reload all data paths based on current data folder setting."""
    global APP_DIR, CONFIG_FILE, REPORTS_FILE, ATTACHMENTS_DIR
    APP_DIR = _get_data_folder()
    CONFIG_FILE = os.path.join(APP_DIR, "config.json")
    REPORTS_FILE = os.path.join(APP_DIR, "reports.json")
    ATTACHMENTS_DIR = os.path.join(APP_DIR, "attachments")

# Initialize paths
APP_DIR = _get_data_folder()
CONFIG_FILE = os.path.join(APP_DIR, "config.json")
REPORTS_FILE = os.path.join(APP_DIR, "reports.json")
ATTACHMENTS_DIR = os.path.join(APP_DIR, "attachments")
