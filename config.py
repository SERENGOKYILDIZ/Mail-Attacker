# config.py

APP_NAME = "MailFlow"
APP_VERSION = "v1.1.0"

# Window settings
WINDOW_WIDTH = 1100
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
APP_DIR = os.path.join(os.environ.get("APPDATA"), APP_NAME)
if not os.path.exists(APP_DIR):
    os.makedirs(APP_DIR)

CONFIG_FILE = os.path.join(APP_DIR, "config.json")
ATTACHMENTS_DIR = os.path.join(APP_DIR, "attachments")
