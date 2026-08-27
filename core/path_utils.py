"""
Path Utilities for KENZEN SeaArt Helper v5.0.0
Resolves application root directory reliably in both source development mode
and PyInstaller compiled executable mode (sys.frozen).
"""

import os
import sys


def get_app_root_dir() -> str:
    """
    Returns the absolute path to the application root directory.
    - If running as a compiled PyInstaller EXE (sys.frozen), returns the folder containing the EXE.
    - If running from source code, returns the workspace directory containing 'kenzen_v5'.
    """
    if getattr(sys, "frozen", False):
        # PyInstaller: sys.executable is the .exe file path
        return os.path.dirname(os.path.abspath(sys.executable))
    
    # Running from Python source: locate the root directory above 'kenzen_v5'
    current_dir = os.path.dirname(os.path.abspath(__file__)) # .../kenzen_v5/core
    kenzen_pkg_dir = os.path.dirname(current_dir)             # .../kenzen_v5
    return os.path.dirname(kenzen_pkg_dir)                   # .../ (Workspace Root)


def get_resource_path(filename: str) -> str:
    """Returns absolute path to a resource file (e.g. 'tags.db', 'KENZEN_Config.json', 'kenzen_icon.ico')."""
    root = get_app_root_dir()
    return os.path.join(root, filename)
