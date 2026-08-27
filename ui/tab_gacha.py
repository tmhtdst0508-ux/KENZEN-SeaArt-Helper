"""
Gacha! Tab for KENZEN SeaArt Helper v5.0.0
AI-powered prompt generator using Google Gemini API (gemini-3.7-flash),
reset-by-default Surprise Me (SFW, NSFW, Hardcore), countdown quota tracking (n / max runs),
double-click quota configurator, 15-second cooldown lock, and input clear controls.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QTextEdit, QGroupBox, QCheckBox, QMessageBox, QApplication, QComboBox,
    QInputDialog
)
from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QMouseEvent, QCursor
from ..core.gemini_api import GeminiAPI
from ..core.config_manager import ConfigManager
from ..core.prompt_engine import sanitize_sd_prompt, sanitize_api_key
from .style import COLOR_ACTION, COLOR_SUCCESS, COLOR_DANGER, safe_copy_to_clipboard


class ClickableStatusLabel(QLabel):
    """Custom label that triggers custom quota configuration dialog on double-click."""
    double_clicked = Signal()

    def __init__(self, text: str = "", parent=None):
        super().__init__(text, parent)
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.setToolTip("💡 Double-click to configure Daily Quota Limit (課金ユーザー向け上限変更)")

    def mouseDoubleClickEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self.double_clicked.emit()
        super().mouseDoubleClickEvent(event)


class TabGacha(QWidget):
    send_to_cockpit = Signal(str)
    send_to_fav = Signal(str, str) # prompt, description
    cooldown_active = Signal(bool)  # (is_locked: bool)

    def __init__(self, gemini_api: GeminiAPI, config_manager: ConfigManager, parent=None):
        super().__init__(parent)
        self.gemini = gemini_api
        self.config = config_manager
        
        self.is_generating = False
        self.cooldown_remaining = 0
        self.cooldown_timer = QTimer(self)
        self.cooldown_timer.setInterval(1000)
        self.cooldown_timer.timeout.connect(self.on_cooldown_tick)

        self.init_ui()

    def is_busy(self) -> bool:
        """Returns True if Gemini API communication is currently in progress."""
        return self.is_generating

    def get_cooldown_remaining(self) -> int:
        """Returns remaining cooldown seconds."""
        return self.cooldown_remaining

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. API Key & Quota Status Bar
        api_box = QGroupBox("Google Gemini API Configuration / API設定")
        api_layout = QHBoxLayout(api_box)
        api_layout.setContentsMargins(10, 8, 10, 8)
        api_layout.setSpacing(8)
        
        lbl_key = QLabel("API Key:")
        lbl_key.setStyleSheet("font-weight: bold;")
        self.txt_api_key = QLineEdit()
        self.txt_api_key.setEchoMode(QLineEdit.Password)
        self.txt_api_key.setPlaceholderText("Enter Gemini API Key...")
        self.txt_api_key.setText(self.config.get_setting("GeminiAPIKey", ""))
        self.txt_api_key.textChanged.connect(self.on_api_key_changed)

        btn_toggle_echo = QPushButton("Show")
        btn_toggle_echo.setFixedWidth(65)
        btn_toggle_echo.clicked.connect(self.toggle_echo_mode)

        # Clickable Quota Status Label (Double-click to change quota limit)
        self.lbl_gacha_status = ClickableStatusLabel()
        self.lbl_gacha_status.setStyleSheet("color: #1E3A8A; font-weight: 600; font-size: 12px; margin-left: 12px;")
        self.lbl_gacha_status.double_clicked.connect(self.on_change_quota_limit)
        self.update_quota_label()

        api_layout.addWidget(lbl_key)
        api_layout.addWidget(self.txt_api_key, 1) # Expand naturally to eliminate dead space
        api_layout.addWidget(btn_toggle_echo)
        api_layout.addWidget(self.lbl_gacha_status)
        layout.addWidget(api_box)

        # 2. Input Box & Surprise Me Options (Default OFF upon launch)
        input_box = QGroupBox("Input Concept / シチュエーション・概念入力")
        input_layout = QVBoxLayout(input_box)
        
        self.txt_input = QTextEdit()
        self.txt_input.setFixedHeight(85)
        self.txt_input.setPlaceholderText("Describe your desired scene, characters, outfit, mood, or background (Japanese or English supported)...")
        input_layout.addWidget(self.txt_input)

        # Modifier Options: Always initialize unchecked (OFF) and SFW
        opt_layout = QHBoxLayout()
        self.chk_surprise = QCheckBox("🎲 Surprise Me! (Inject creative random inspiration tags)")
        self.chk_surprise.setChecked(False) # Always start unchecked upon launch
        self.chk_surprise.stateChanged.connect(self.on_surprise_toggled)

        lbl_s_level = QLabel("Surprise Level:")
        self.cmb_surprise_level = QComboBox()
        self.cmb_surprise_level.addItems(["SFW", "NSFW", "Hardcore"])
        self.cmb_surprise_level.setCurrentText("SFW") # Always start with SFW upon launch
        self.cmb_surprise_level.setEnabled(False)

        btn_clear_input = QPushButton("Clear Input")
        btn_clear_input.setProperty("btnType", "danger")
        btn_clear_input.clicked.connect(lambda: self.txt_input.clear())

        self.btn_gacha = QPushButton("✨ Roll Gacha! / ガチャを回す！")
        self.btn_gacha.setProperty("btnType", "success")
        self.btn_gacha.setFixedHeight(36)
        self.btn_gacha.clicked.connect(self.on_spin_gacha)

        opt_layout.addWidget(self.chk_surprise)
        opt_layout.addWidget(lbl_s_level)
        opt_layout.addWidget(self.cmb_surprise_level)
        opt_layout.addStretch()
        opt_layout.addWidget(btn_clear_input)
        opt_layout.addWidget(self.btn_gacha)
        input_layout.addLayout(opt_layout)

        layout.addWidget(input_box)

        # 3. Output Result Box
        res_box = QGroupBox("AI Generated Visual Tags / AI生成結果")
        res_layout = QVBoxLayout(res_box)

        self.txt_result = QTextEdit()
        self.txt_result.setPlaceholderText("Generated visual descriptive tags will appear here...")
        res_layout.addWidget(self.txt_result)

        # Result Action Sub-bar
        res_btn_bar = QHBoxLayout()
        self.btn_clear_res = QPushButton("Clear Result")
        self.btn_clear_res.setProperty("btnType", "danger")
        self.btn_clear_res.clicked.connect(lambda: self.txt_result.clear())

        res_btn_bar.addStretch()
        res_btn_bar.addWidget(self.btn_clear_res)
        res_layout.addLayout(res_btn_bar)

        layout.addWidget(res_box, 1)

        # 4. Bottom Action Bar
        bottom_bar = QHBoxLayout()

        self.btn_send_cockpit = QPushButton("🚀 Send to Cockpit")
        self.btn_send_cockpit.setProperty("btnType", "success")
        self.btn_send_cockpit.setFixedHeight(38)
        self.btn_send_cockpit.clicked.connect(self.on_send_cockpit)

        self.btn_send_fav = QPushButton("⭐ Send to Favorites / お気に入り転送")
        self.btn_send_fav.setProperty("btnType", "action")
        self.btn_send_fav.setFixedHeight(38)
        self.btn_send_fav.clicked.connect(self.on_send_fav)

        self.btn_copy = QPushButton("📋 Copy to Clipboard")
        self.btn_copy.setFixedHeight(38)
        self.btn_copy.clicked.connect(self.on_copy)

        bottom_bar.addWidget(self.btn_send_cockpit, 2)
        bottom_bar.addWidget(self.btn_send_fav, 1)
        bottom_bar.addWidget(self.btn_copy, 1)
        layout.addLayout(bottom_bar)

    def on_surprise_toggled(self, state: int):
        is_checked = bool(state)
        self.cmb_surprise_level.setEnabled(is_checked)

    def toggle_echo_mode(self):
        if self.txt_api_key.echoMode() == QLineEdit.Password:
            self.txt_api_key.setEchoMode(QLineEdit.Normal)
        else:
            self.txt_api_key.setEchoMode(QLineEdit.Password)

    def on_api_key_changed(self, text: str):
        key = sanitize_api_key(text)
        self.gemini.set_api_key(key)
        self.config.set_setting("GeminiAPIKey", key)

    def update_quota_label(self):
        pt_date, remaining, max_quota = self.config.get_gacha_status_counts()
        self.lbl_gacha_status.setText(f"Daily Runs: {remaining} / {max_quota} (PT: {pt_date})")

    def on_change_quota_limit(self):
        """Double-click gimmick allowing users to change their daily quota limit."""
        current_max = self.config.get_gacha_quota()
        new_val, ok = QInputDialog.getInt(
            self,
            "上限回数の変更 / Set Daily Quota Limit",
            "1日の上限回数（Daily Quota Limit）を入力してください:\n(無料プラン: 15回 / 課金プラン: 1500回等)",
            current_max,
            1,
            99999,
            1
        )
        if ok and new_val > 0:
            self.config.set_gacha_quota(new_val)
            self.update_quota_label()
            QMessageBox.information(
                self,
                "設定完了 / Success",
                f"1日の上限回数を {new_val} 回に設定しました。\nDaily quota limit set to {new_val} runs."
            )

    def on_spin_gacha(self):
        if self.cooldown_remaining > 0:
            return

        pt_date, remaining, max_quota = self.config.get_gacha_status_counts()
        if remaining <= 0:
            QMessageBox.warning(
                self,
                "回数上限 / Quota Limit Exceeded",
                f"本日の無料生成枠（{max_quota}回）を使い切りました。\n"
                f"PT 0:00（日本時間 16:00/17:00）にリセットされます。\n\n"
                f"You have reached today's limit of {max_quota} runs."
            )
            return

        # 1. Check API Key First
        api_key = sanitize_api_key(self.txt_api_key.text().strip())
        if not api_key:
            QMessageBox.warning(
                self,
                "APIキー未入力 / Missing API Key",
                "Google Gemini API Key を入力してください。\nPlease provide a valid Gemini API Key."
            )
            return

        # 2. Check Text Input or Surprise Me
        user_input = self.txt_input.toPlainText().strip()
        is_surprise = self.chk_surprise.isChecked()

        if not user_input and not is_surprise:
            QMessageBox.warning(
                self,
                "入力不足 / Input Required",
                "テキストを入力するか、Surprise Me! を有効にしてください。\nPlease enter concept text or enable Surprise Me!."
            )
            return

        surprise_tags = None
        if is_surprise:
            s_level = self.cmb_surprise_level.currentText()
            surprise_tags = self.gemini.get_surprise_keywords(s_level)
            if surprise_tags:
                self.txt_result.setPlainText(
                    f"🎲 [Surprise Me! 抽出タグ / Injected Inspiration Tags ({s_level})]\n"
                    f"{surprise_tags}\n\n"
                    f"🤖 [Gemini 思考中 / Generating with gemini-3.7-flash...]"
                )
            else:
                self.txt_result.setPlainText("🤖 [Gemini 思考中 / Generating with gemini-3.7-flash...]")
            
            if not user_input:
                user_input = f"Random aesthetic scene with tags: {surprise_tags}"
        else:
            self.txt_result.setPlainText("🤖 [Gemini 思考中 / Generating with gemini-3.7-flash...]")

        self.is_generating = True
        self.btn_gacha.setEnabled(False)
        self.btn_gacha.setText("🤖 Generating / 生成中...")
        QApplication.processEvents()

        try:
            self.gemini.set_api_key(api_key)
            success, result = self.gemini.generate_gacha_prompt(
                user_input,
                precomputed_surprise_tags=surprise_tags
            )

            if success:
                cleaned_result = sanitize_sd_prompt(result)
                self.txt_result.setPlainText(cleaned_result)
                # Deduct 1 run only upon success
                used = self.config.get_setting("GachaCount", 0) + 1
                self.config.set_setting("GachaCount", used)
                self.config.set_setting("LastPTDate", GeminiAPI.get_pt_date_str())
                self.update_quota_label()
                
                self.start_cooldown(15)
            else:
                QMessageBox.critical(
                    self,
                    "生成エラー / Generation Error",
                    result
                )
        finally:
            self.is_generating = False
            if self.cooldown_remaining <= 0:
                self.btn_gacha.setEnabled(True)
                self.btn_gacha.setText("✨ Roll Gacha! / ガチャを回す！")

    def start_cooldown(self, seconds: int = 15):
        self.cooldown_remaining = seconds
        self.btn_gacha.setEnabled(False)
        self.btn_send_cockpit.setEnabled(False)
        self.btn_send_fav.setEnabled(False)
        self.cooldown_active.emit(True)
        self.cooldown_timer.start()
        self.on_cooldown_tick()

    def on_cooldown_tick(self):
        if self.cooldown_remaining > 0:
            self.btn_gacha.setText(f"⏳ Cooldown ({self.cooldown_remaining}s)...")
            self.cooldown_remaining -= 1
        else:
            self.cooldown_timer.stop()
            self.btn_gacha.setEnabled(True)
            self.btn_send_cockpit.setEnabled(True)
            self.btn_send_fav.setEnabled(True)
            self.btn_gacha.setText("✨ Roll Gacha! / ガチャを回す！")
            self.cooldown_active.emit(False)

    def on_send_cockpit(self):
        res = self.txt_result.toPlainText().strip()
        if res and not res.startswith("🤖 [Gemini 思考中") and not res.startswith("🎲 [Surprise Me!"):
            cleaned = sanitize_sd_prompt(res)
            self.send_to_cockpit.emit(cleaned)
        else:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "送信可能な生成結果がありません。\nResult is empty."
            )

    def on_send_fav(self):
        res = self.txt_result.toPlainText().strip()
        desc = self.txt_input.toPlainText().strip() or "Gacha Generated"
        if res and not res.startswith("🤖 [Gemini 思考中") and not res.startswith("🎲 [Surprise Me!"):
            cleaned = sanitize_sd_prompt(res)
            self.send_to_fav.emit(cleaned, desc[:50])
        else:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "送信可能な生成結果がありません。\nResult is empty."
            )

    def on_copy(self):
        res = self.txt_result.toPlainText().strip()
        if not res or res.startswith("🤖 [Gemini 思考中") or res.startswith("🎲 [Surprise Me!"):
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "生成結果が空です。コピーする内容がありません。\nResult is empty."
            )
            return

        cleaned = sanitize_sd_prompt(res)
        safe_copy_to_clipboard(cleaned, self)
        QMessageBox.information(
            self,
            "完了 / Success",
            "クリップボードにコピーしました！\nCopied to clipboard!"
        )

    def clear_all_ui(self):
        """Clears all Gacha UI inputs and resets controls on NUKE."""
        self.txt_api_key.clear()
        self.gemini.set_api_key("")
        self.txt_input.clear()
        self.txt_result.clear()
        self.chk_surprise.setChecked(False)
        self.cmb_surprise_level.setCurrentText("SFW")
        self.cmb_surprise_level.setEnabled(False)
