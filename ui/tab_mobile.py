"""
Mobile Memo Tab for KENZEN SeaArt Helper v5.0.0
Manages quick mobile memos ([Memo], [URL], [Hash] badge on Left list, Content-only on Right edit),
clipboard copy, Mobile Memo JSON import, and synchronizing with Master Config.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QPushButton,
    QListWidget, QGroupBox, QSplitter, QMessageBox, QApplication, QFileDialog,
    QListWidgetItem
)
from PySide6.QtCore import Qt, Signal
from ..core.config_manager import ConfigManager, parse_mobile_memo_item, format_mobile_memo_item
from .style import COLOR_ACTION, COLOR_SUCCESS, COLOR_DANGER
from .widgets import PlainTextOnlyTextEdit


class TabMobile(QWidget):
    def __init__(self, config_manager: ConfigManager, parent=None):
        super().__init__(parent)
        self.config = config_manager
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        splitter = QSplitter(Qt.Horizontal)

        # 1. Left: Memo List (Shows [Memo], [URL], [Hash] badges)
        left_box = QGroupBox("Mobile Memo List / メモ一覧")
        left_layout = QVBoxLayout(left_box)

        self.memo_list = QListWidget()
        self.memo_list.itemClicked.connect(self.on_memo_selected)
        self.memo_list.itemDoubleClicked.connect(self.on_memo_double_clicked)
        left_layout.addWidget(self.memo_list)

        # Delete key shortcut on memo_list (Issue 2)
        from PySide6.QtGui import QKeySequence, QShortcut
        self.shortcut_del = QShortcut(QKeySequence.Delete, self.memo_list)
        self.shortcut_del.activated.connect(self.on_delete_memo)

        list_btn_bar = QHBoxLayout()
        btn_del = QPushButton("Delete Selected")
        btn_del.setProperty("btnType", "danger")
        btn_del.clicked.connect(self.on_delete_memo)

        btn_clear_all = QPushButton("Clear All")
        btn_clear_all.setProperty("btnType", "danger")
        btn_clear_all.clicked.connect(self.on_clear_all)

        list_btn_bar.addWidget(btn_del)
        list_btn_bar.addWidget(btn_clear_all)
        left_layout.addLayout(list_btn_bar)

        splitter.addWidget(left_box)

        # 2. Right: Detail / Edit (Issue 5: Content only)
        right_box = QGroupBox("Memo Content / メモ内容・編集")
        right_layout = QVBoxLayout(right_box)

        self.txt_content = PlainTextOnlyTextEdit()
        self.txt_content.setPlaceholderText("Enter or edit content (prompts, URL, or notes)...")
        right_layout.addWidget(self.txt_content)

        right_btn_bar = QHBoxLayout()
        btn_add_new = QPushButton("➕ Add as New Memo")
        btn_add_new.clicked.connect(self.on_add_memo)

        btn_update = QPushButton("💾 Update Selected")
        btn_update.clicked.connect(self.on_update_memo)

        right_btn_bar.addWidget(btn_add_new)
        right_btn_bar.addWidget(btn_update)
        right_layout.addLayout(right_btn_bar)

        splitter.addWidget(right_box)
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 3)

        layout.addWidget(splitter, 1)

        # 3. Bottom Actions
        bottom_bar = QHBoxLayout()

        btn_copy = QPushButton("📋 Copy to Clipboard")
        btn_copy.setProperty("btnType", "success")
        btn_copy.setFixedHeight(38)
        btn_copy.clicked.connect(self.on_copy)

        btn_import_mobile = QPushButton("📱 Import Mobile Memos from JSON (モバイルメモ取込)")
        btn_import_mobile.setProperty("btnType", "action")
        btn_import_mobile.setFixedHeight(38)
        btn_import_mobile.clicked.connect(self.on_import_mobile_json)

        btn_save_config = QPushButton("💾 Save Memos to Config (母艦に保存)")
        btn_save_config.setFixedHeight(38)
        btn_save_config.clicked.connect(self.on_save_to_config)

        bottom_bar.addWidget(btn_copy, 1)
        bottom_bar.addWidget(btn_import_mobile, 2)
        bottom_bar.addWidget(btn_save_config, 1)
        layout.addLayout(bottom_bar)

        self.load_memos()

    def load_memos(self):
        """Populates mobile memos with [Memo], [URL], [Hash] badges on Left."""
        memos = self.config.get_mobile_memos()
        self.memo_list.clear()
        for idx, m in enumerate(memos):
            m_type, m_content = parse_mobile_memo_item(m)
            first_line = m_content.strip().split("\n")[0] if m_content.strip() else "(Empty)"
            item = QListWidgetItem(f"{idx+1}. [{m_type}] {first_line[:35]}")
            item.setData(Qt.UserRole, (m_type, m_content))
            self.memo_list.addItem(item)

    def on_memo_selected(self, item: QListWidgetItem):
        """Issue 5: Displays pure content on the Right edit box."""
        data = item.data(Qt.UserRole)
        if data:
            m_type, m_content = data
            self.txt_content.setPlainText(m_content)
        else:
            row = self.memo_list.row(item)
            memos = self.config.get_mobile_memos()
            if 0 <= row < len(memos):
                _, m_content = parse_mobile_memo_item(memos[row])
                self.txt_content.setPlainText(m_content)

    def on_memo_double_clicked(self, item: QListWidgetItem):
        """
        Issue 2: Double click action:
        - If [URL]: Open in default web browser.
        - If [Hash]: Copy to clipboard and notify.
        - If [Memo]: Focus edit box.
        """
        data = item.data(Qt.UserRole)
        if not data:
            return
        m_type, m_content = data
        m_type_upper = m_type.upper()

        if m_type_upper == "URL":
            from PySide6.QtGui import QDesktopServices
            from PySide6.QtCore import QUrl
            url_str = m_content.strip()
            if not (url_str.startswith("http://") or url_str.startswith("https://")):
                url_str = f"https://{url_str}"
            QDesktopServices.openUrl(QUrl(url_str))
        elif m_type_upper == "HASH":
            QApplication.clipboard().setText(m_content.strip())
            QMessageBox.information(
                self,
                "コピー完了 / Success",
                f"ハッシュ値をクリップボードにコピーしました！\nHash copied to clipboard:\n{m_content}"
            )
        else:
            self.txt_content.setFocus()

    def on_add_memo(self):
        content = self.txt_content.toPlainText().strip()
        if not content:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "内容を入力してください。\nPlease enter memo content."
            )
            return

        m_type, clean_content = parse_mobile_memo_item(content)
        formatted = f"[{m_type}] {clean_content}"
        
        memos = self.config.get_mobile_memos()
        memos.append(formatted)
        self.config.set_mobile_memos(memos)
        self.load_memos()
        self.memo_list.setCurrentRow(len(memos) - 1)
        QMessageBox.information(
            self,
            "登録完了 / Success",
            f"[{m_type}] メモを登録しました！\nMemo added!"
        )

    def on_update_memo(self):
        row = self.memo_list.currentRow()
        if row < 0:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "更新するメモを選択してください。\nPlease select a memo to update."
            )
            return

        content = self.txt_content.toPlainText().strip()
        if not content:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "内容を入力してください。\nPlease enter memo content."
            )
            return

        m_type, clean_content = parse_mobile_memo_item(content)
        formatted = f"[{m_type}] {clean_content}"

        memos = self.config.get_mobile_memos()
        if row < len(memos):
            memos[row] = formatted
            self.config.set_mobile_memos(memos)
            self.load_memos()
            self.memo_list.setCurrentRow(row)
            QMessageBox.information(
                self,
                "保存完了 / Success",
                f"[{m_type}] メモを更新しました。\nMemo updated."
            )

    def on_delete_memo(self):
        row = self.memo_list.currentRow()
        if row < 0:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "削除するメモを選択してください。\nPlease select a memo to delete."
            )
            return
        memos = self.config.get_mobile_memos()
        if row < len(memos):
            memos.pop(row)
            self.config.set_mobile_memos(memos)
            self.load_memos()
            self.txt_content.clear()

    def on_clear_all(self):
        memos = self.config.get_mobile_memos()
        if not memos:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "メモリストは既に空です。\nMemo list is already empty."
            )
            return

        ans = QMessageBox.question(
            self,
            "削除確認 / Confirm Delete",
            "すべてのメモを削除しますか？\nAre you sure you want to delete all memos?"
        )
        if ans == QMessageBox.Yes:
            self.config.set_mobile_memos([])
            self.load_memos()
            self.txt_content.clear()

    def on_copy(self):
        """Copies content of the selected memo or edit box."""
        content = self.txt_content.toPlainText().strip()
        if not content:
            row = self.memo_list.currentRow()
            if row >= 0:
                item = self.memo_list.item(row)
                if item and item.data(Qt.UserRole):
                    content = item.data(Qt.UserRole)[1]

        if content:
            QApplication.clipboard().setText(content)
            QMessageBox.information(
                self,
                "完了 / Success",
                "クリップボードにコピーしました！\nCopied to clipboard!"
            )
        else:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "コピーするメモが選択されていません。\nNo memo content to copy."
            )

    def on_import_mobile_json(self):
        fname, _ = QFileDialog.getOpenFileName(
            self,
            "モバイル用JSONの選択 / Select Mobile JSON",
            "",
            "JSON Files (*.json)"
        )
        if fname:
            if self.config.is_master_config_file(fname):
                QMessageBox.critical(
                    self,
                    "インポートエラー / Import Error",
                    "「KENZEN_Config.json」は母艦設定ファイルのため、モバイルメモのインポート対象外です。\n\n"
                    "モバイル用JSON（KENZEN_Mobile_Fav.json や KENZEN_Mobile.json）を選択してください。\n\n"
                    "'KENZEN_Config.json' cannot be imported as a Mobile Memo file."
                )
                return

            try:
                added_memos = self.config.import_mobile_memos_only(fname)
                self.load_memos()
                QMessageBox.information(
                    self,
                    "インポート完了 / Success",
                    f"モバイルメモをインポートしました！\nImported {added_memos} Mobile Memo(s)."
                )
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to import mobile JSON: {e}")

    def on_save_to_config(self):
        memos = self.config.get_mobile_memos()
        self.config.set_mobile_memos(memos)
        QMessageBox.information(
            self,
            "同期完了 / Success",
            f"母艦のコンフィグファイルに {len(memos)} 件のメモを保存しました。\nMemos successfully saved to Master Config."
        )
