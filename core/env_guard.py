"""
Environment Guard for KENZEN SeaArt Helper v5.0.0
Protects user prompt assets by detecting execution within Cloud Sync Folders
(OneDrive, Google Drive, Dropbox, iCloud, Box, Nextcloud, etc.)
and strictly halting launch before any SQLite / JSON file I/O occurs.
"""

import os
import sys
from typing import Tuple, Optional


CLOUD_FOLDER_SIGNATURES = [
    "onedrive",
    "google drive",
    "googledrive",
    "マイドライブ",
    "dropbox",
    "iclouddrive",
    "box sync",
    "boxsync",
    "nextcloud",
    "owncloud"
]

CLOUD_ENV_VARS = [
    "OneDrive",
    "OneDriveConsumer",
    "OneDriveCommercial",
    "Dropbox",
    "GoogleDriveFS"
]


def check_cloud_sync_environment(target_path: Optional[str] = None) -> Tuple[bool, str]:
    """
    Checks if the given path (or current working directory / script directory)
    is located inside a cloud sync folder.
    Returns (is_cloud: bool, reason_message: str).
    """
    if not target_path:
        target_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    norm_target = os.path.normpath(os.path.abspath(target_path)).lower()

    # 1. Check Cloud Environment Variables
    for env_var in CLOUD_ENV_VARS:
        env_val = os.environ.get(env_var)
        if env_val:
            norm_env = os.path.normpath(os.path.abspath(env_val)).lower()
            if norm_target.startswith(norm_env) or norm_env in norm_target:
                return True, f"検出された環境変数: {env_var} ({env_val})"

    # 2. Check Path String Signatures
    # Normalize separators to '/' for universal parsing
    target_clean = norm_target.replace("\\", "/")
    path_components = [c.strip() for c in target_clean.split("/") if c.strip()]

    for sig in CLOUD_FOLDER_SIGNATURES:
        for comp in path_components:
            # Matches 'onedrive', 'onedrive - personal', 'onedrive - company', 'dropbox (personal)', etc.
            if comp == sig or comp.startswith(f"{sig} ") or comp.startswith(f"{sig}-") or comp.startswith(f"{sig}_"):
                return True, f"検出されたクラウドフォルダ名: '{comp}'"

    # 3. Check for WebDAV / URL paths
    if norm_target.startswith("http:") or norm_target.startswith("https:") or norm_target.startswith("\\\\?\\http"):
        return True, "検出されたネットワーク/WebDAVパス"

    return False, ""


def enforce_local_execution_or_exit(app_root_path: Optional[str] = None):
    """
    Strictly checks the execution environment.
    If inside a cloud sync folder, shows a critical dialog and terminates the process immediately.
    """
    if not app_root_path:
        app_root_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    is_cloud, detail = check_cloud_sync_environment(app_root_path)
    if not is_cloud:
        return # Safe to proceed

    # If Qt application is available or need to initialize for dialog
    from PySide6.QtWidgets import QApplication, QMessageBox

    app = QApplication.instance() or QApplication(sys.argv)

    msg = (
        "【⚠️ 起動中止 / Launch Blocked】\n"
        "クラウド同期フォルダ上では動作できません。\n"
        "Cannot run inside a cloud sync folder.\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "【理由 / Reason】\n"
        "同期の競合によるデータ破損・プロンプト消失を防ぐためです。\n"
        "To protect your prompt assets from sync conflicts & corruption.\n\n"
        "【対処方法 / Action Required】\n"
        "フォルダ全体をローカルドライブ（D:\\ や C:\\ 直下等）へ移動してください。\n"
        "Please move this entire folder to a local drive (e.g. D:\\ or C:\\).\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        f"Path: {app_root_path}\n"
        f"Detail: {detail}"
    )

    box = QMessageBox()
    box.setIcon(QMessageBox.Critical)
    box.setWindowTitle("重大警告 / Cloud Sync Detected")
    box.setText(msg)
    box.setStandardButtons(QMessageBox.Ok)
    box.exec()

    sys.exit(1)
