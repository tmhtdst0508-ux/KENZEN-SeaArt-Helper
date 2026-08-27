"""
Style definitions and Theme configurations for KENZEN SeaArt Helper v5.0.0
Faithfully recreates the VBA color schemes and visual hierarchy.
"""

# Color Constants
COLOR_ACTION = "#C8E1FF"       # Sky Blue (VBA colAction: RGB(200, 225, 255))
COLOR_ACTION_HOVER = "#B0D4FF"
COLOR_ACTION_PRESSED = "#98C6FF"

COLOR_SUCCESS = "#C8F0C8"      # Light Green (VBA colSuccess: RGB(200, 240, 200))
COLOR_SUCCESS_HOVER = "#B0E8B0"
COLOR_SUCCESS_PRESSED = "#98DE98"

COLOR_DANGER = "#FFD2D2"       # Light Red (Clear/Delete actions)
COLOR_DANGER_HOVER = "#FFB8B8"
COLOR_DANGER_PRESSED = "#FFA0A0"

COLOR_WARNING = "#FEF08A"      # Light Amber/Yellow (Notice/History preserved clear)
COLOR_WARNING_HOVER = "#FDE047"
COLOR_WARNING_PRESSED = "#EAB308"

COLOR_NEUTRAL = "#F0F0F0"      # Light Gray button default
COLOR_NEUTRAL_HOVER = "#E4E4E4"
COLOR_NEUTRAL_PRESSED = "#D8D8D8"

COLOR_DISABLED_BG = "#E0E0E0"
COLOR_DISABLED_FG = "#808080"

# Tag Matrix colors
COLOR_TAG_EN_BG = "#BAE6FD"    # Light Cyan for English prompt button
COLOR_TAG_EN_HOVER = "#93C5FD"
COLOR_TAG_JP_BG = "#F8FAFC"    # Soft off-white for Japanese label
COLOR_CATEGORY_HEADER = "#334155" # Slate dark for category titles
COLOR_CATEGORY_TEXT = "#FFFFFF"

MAIN_STYLESHEET = f"""
QMainWindow {{
    background-color: #F8F9FA;
}}

QTabWidget::pane {{
    border: 1px solid #D1D5DB;
    background-color: #FFFFFF;
    border-radius: 4px;
}}

QTabBar::tab {{
    background-color: #E5E7EB;
    color: #1F2937;
    padding: 8px 14px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    font-weight: 500;
    font-size: 13px;
}}

QTabBar::tab:selected {{
    background-color: #FFFFFF;
    color: #1E40AF;
    font-weight: bold;
    border-top: 2px solid #2563EB;
}}

QTabBar::tab:hover:!selected {{
    background-color: #D1D5DB;
}}

QGroupBox {{
    font-weight: bold;
    border: 1px solid #D1D5DB;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 12px;
    background-color: #FFFFFF;
}}

QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 4px;
    color: #374151;
}}

QLineEdit, QTextEdit, QPlainTextEdit {{
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 6px;
    color: #1E293B;
    font-size: 13px;
    selection-background-color: #93C5FD;
}}

QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {{
    border: 1px solid #3B82F6;
}}

QComboBox {{
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 4px 8px;
    font-size: 13px;
    color: #1E293B;
}}

QComboBox:hover {{
    border: 1px solid #94A3B8;
}}

QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 20px;
    border-left: 1px solid #E2E8F0;
}}

QListWidget {{
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 4px;
    font-size: 13px;
}}

QListWidget::item {{
    padding: 4px 8px;
    border-radius: 3px;
}}

QListWidget::item:selected {{
    background-color: #BFDBFE;
    color: #1E3A8A;
}}

QListWidget::item:hover:!selected {{
    background-color: #F1F5F9;
}}

QPushButton {{
    background-color: {COLOR_NEUTRAL};
    border: 1px solid #CBD5E1;
    border-radius: 4px;
    padding: 6px 12px;
    font-size: 13px;
    color: #1E293B;
}}

QPushButton:hover {{
    background-color: {COLOR_NEUTRAL_HOVER};
}}

QPushButton:pressed {{
    background-color: {COLOR_NEUTRAL_PRESSED};
}}

QPushButton:disabled {{
    background-color: {COLOR_DISABLED_BG};
    color: {COLOR_DISABLED_FG};
    border-color: #D1D5DB;
}}

QPushButton[btnType="action"] {{
    background-color: {COLOR_ACTION};
    color: #0F172A;
    font-weight: 500;
}}

QPushButton[btnType="action"]:hover {{
    background-color: {COLOR_ACTION_HOVER};
}}

QPushButton[btnType="action"]:pressed {{
    background-color: {COLOR_ACTION_PRESSED};
}}

QPushButton[btnType="success"] {{
    background-color: {COLOR_SUCCESS};
    color: #064E3B;
    font-weight: bold;
}}

QPushButton[btnType="success"]:hover {{
    background-color: {COLOR_SUCCESS_HOVER};
}}

QPushButton[btnType="success"]:pressed {{
    background-color: {COLOR_SUCCESS_PRESSED};
}}

QPushButton[btnType="danger"] {{
    background-color: {COLOR_DANGER};
    color: #7F1D1D;
}}

QPushButton[btnType="danger"]:hover {{
    background-color: {COLOR_DANGER_HOVER};
}}

QPushButton[btnType="danger"]:pressed {{
    background-color: {COLOR_DANGER_PRESSED};
}}

QPushButton[btnType="warning"] {{
    background-color: {COLOR_WARNING};
    color: #78350F;
    font-weight: 500;
}}

QPushButton[btnType="warning"]:hover {{
    background-color: {COLOR_WARNING_HOVER};
}}

QPushButton[btnType="warning"]:pressed {{
    background-color: {COLOR_WARNING_PRESSED};
}}

QCheckBox {{
    font-size: 13px;
    color: #374151;
    spacing: 6px;
}}

QStatusBar {{
    background-color: #F1F5F9;
    color: #475569;
    border-top: 1px solid #E2E8F0;
}}
"""


def safe_copy_to_clipboard(text: str, parent=None) -> bool:
    """
    Edge-case Guard 5: Safely copies text to clipboard with retries,
    catching Windows clipboard lock/contention errors gracefully.
    """
    import time
    from PySide6.QtWidgets import QApplication, QMessageBox

    if not text:
        return False

    for attempt in range(3):
        try:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            return True
        except Exception as e:
            time.sleep(0.05)

    try:
        if parent:
            QMessageBox.warning(
                parent,
                "クリップボード警告 / Clipboard Warning",
                "他の常駐アプリがクリップボードを使用中のため、コピーに失敗しました。\nFailed to access system clipboard."
            )
    except Exception:
        pass
    return False
