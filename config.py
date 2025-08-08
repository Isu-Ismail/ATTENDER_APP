import os
import sys

# --- Configuration ---
ICON_PATH = "app.ico"

# --- Smart Path Definitions ---

def resource_path(relative_path):
    """ Get absolute path to a bundled resource, works for dev and for PyInstaller. """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Defines the stable folder in the user's Documents for saving .xlsx files.
USER_DATA_PATH = os.path.join(os.path.expanduser('~'), 'Documents', 'AttendanceMarker')

# Creates this folder automatically if it doesn't exist.
os.makedirs(USER_DATA_PATH, exist_ok=True)