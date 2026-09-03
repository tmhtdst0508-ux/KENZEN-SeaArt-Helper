"""
Main Entry Point for KENZEN SeaArt Helper v5.0.1
"""

import sys
import os

# Add parent directory to sys.path if needed
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon

try:
    from kenzen_v5.core.env_guard import enforce_local_execution_or_exit
    from kenzen_v5.core.single_instance import SingleInstanceManager
    from kenzen_v5.core.path_utils import get_app_root_dir, get_resource_path
    from kenzen_v5.ui.main_window import MainWindow
except ImportError:
    from core.env_guard import enforce_local_execution_or_exit
    from core.single_instance import SingleInstanceManager
    from core.path_utils import get_app_root_dir, get_resource_path
    from ui.main_window import MainWindow


def main():
    root_dir = get_app_root_dir()

    # 1. Edge-case Guard: Strictly prevent execution within cloud sync folders to protect prompt assets
    enforce_local_execution_or_exit(root_dir)

    # Windows Taskbar Icon Fix: Register unique AppUserModelID
    if sys.platform == "win32":
        try:
            import ctypes
            myappid = "kenzen.seaart.helper.v5.0.1"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

    # Enable high DPI scaling
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    app = QApplication(sys.argv)
    app.setApplicationName("KENZEN SeaArt Helper v5.0.1")

    # 2. Edge-case Guard: Prevent multiple instances, notify active instance, and bring to front
    single_inst = SingleInstanceManager(app)
    if single_inst.check_already_running():
        sys.exit(0)

    # Set Application Icon
    icon_path = get_resource_path("kenzen_icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    single_inst.activate_requested.connect(window.bring_to_front)
    app.aboutToQuit.connect(single_inst.cleanup)

    # Simultaneously show both Main Window and Matrix Window upon start
    window.show_windows_side_by_side()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
