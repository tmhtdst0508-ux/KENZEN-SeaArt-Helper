"""
Cockpit Tab for KENZEN SeaArt Helper v5.2.0
Main cockpit for prompt editing, DeepL real-time translation assistant,
weighting, Dynamic Prompts wrapping ({A | B | C}),
category sorting with Base Positive & LoRA triggers, two clear buttons (Soft & Hard),
and clipboard operations.
"""

import re
from typing import Optional
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QLineEdit,
    QPushButton, QComboBox, QCheckBox, QGroupBox, QMessageBox, QApplication,
    QDialog, QGridLayout
)
from PySide6.QtCore import Qt, Signal, QUrl
from PySide6.QtGui import QTextCursor, QDesktopServices
from ..core.prompt_engine import PromptEngine, sanitize_sd_prompt
from ..core.config_manager import ConfigManager
from ..core.db_manager import DBManager
from ..core.deepl_api import DeepLAPI, DeepLTranslateWorker
from .style import COLOR_ACTION, COLOR_SUCCESS, COLOR_DANGER, COLOR_WARNING
from .widgets import PlainTextOnlyTextEdit


class DeepLKeyDialog(QDialog):
    """Simple, clean modal dialog to configure DeepL API key."""
    def __init__(self, config: ConfigManager, deepl_api: DeepLAPI, parent=None):
        super().__init__(parent)
        self.config = config
        self.deepl_api = deepl_api
        self.setWindowTitle("DeepL API Key Configuration / API設定")
        self.setFixedWidth(460)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        lbl_desc = QLabel(
            "<b>Please configure your DeepL API Key (Free or Pro). / DeepL APIキーを設定してください。</b><br>"
            "<span style='color: #475569; font-size: 11px;'>"
            "Supports DeepL API Free keys (ends with <code>:fx</code>, 500,000 chars/month free).<br>"
            "DeepL API Free（月50万文字無料）のキーに対応。キーはWindowsレジストリ内に安全に保存されます。</span>"
        )
        lbl_desc.setWordWrap(True)
        layout.addWidget(lbl_desc)

        h_key = QHBoxLayout()
        self.txt_key = QLineEdit()
        self.txt_key.setEchoMode(QLineEdit.Password)
        self.txt_key.setPlaceholderText("Enter DeepL API Key (e.g. xxxxxxxx-...:fx)")
        self.txt_key.setText(self.config.get_setting("DeepLAPIKey", ""))
        
        btn_show = QPushButton("Show")
        btn_show.setFixedWidth(55)
        btn_show.clicked.connect(self._toggle_show)

        h_key.addWidget(self.txt_key, 1)
        h_key.addWidget(btn_show)
        layout.addLayout(h_key)

        # Direct Universal Link to DeepL API Key management (auto-redirects to localized language)
        btn_link = QPushButton("🔗 Get DeepL API Key / キー取得ページを開く (公式サイト)")
        btn_link.setToolTip("ブラウザでDeepL公式サイトのAPIキー取得ページを開きます (自動で居住地言語にリダイレクト)")
        btn_link.setProperty("btnType", "action")
        btn_link.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(DeepLAPI.KEY_URL)))
        layout.addWidget(btn_link)

        # Action buttons
        btn_bar = QHBoxLayout()
        btn_bar.addStretch()

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        btn_save = QPushButton("Save / 保存")
        btn_save.setProperty("btnType", "success")
        btn_save.clicked.connect(self._on_save)

        btn_bar.addWidget(btn_cancel)
        btn_bar.addWidget(btn_save)
        layout.addLayout(btn_bar)

    def _toggle_show(self):
        if self.txt_key.echoMode() == QLineEdit.Password:
            self.txt_key.setEchoMode(QLineEdit.Normal)
        else:
            self.txt_key.setEchoMode(QLineEdit.Password)

    def _on_save(self):
        key = self.txt_key.text().strip()
        self.config.set_setting("DeepLAPIKey", key)
        self.deepl_api.set_api_key(key)
        self.accept()


class TabCockpit(QWidget):
    open_matrix_requested = Signal()
    send_to_fav_requested = Signal(str, str) # prompt, description
    user_tag_added = Signal(dict)            # Emitted when a new tag is registered to user dictionary

    def __init__(self, prompt_engine: PromptEngine, config_manager: ConfigManager, db_manager: Optional[DBManager] = None, parent=None):
        super().__init__(parent)
        self.engine = prompt_engine
        self.config = config_manager
        self.db = db_manager
        self.deepl_api = DeepLAPI(self.config.get_setting("DeepLAPIKey", ""))
        self._trans_worker: Optional[DeepLTranslateWorker] = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. Top Control Bar (Sleek App Title & Matrix Launcher Button)
        top_bar = QHBoxLayout()
        top_bar.setContentsMargins(2, 0, 2, 0)
        top_bar.setSpacing(12)

        lbl_title = QLabel("🎨 <b style='color: #1E293B; font-size: 13px; font-family: Segoe UI, sans-serif;'>KENZEN SeaArt Helper</b> <span style='color: #475569; font-size: 11px; background-color: #E2E8F0; padding: 2px 6px; border-radius: 4px; font-weight: 600;'>v5.2.0</span>")
        lbl_title.setStyleSheet("padding: 2px;")

        btn_matrix = QPushButton("📖 Show Dictionary Matrix (辞書マトリクス表示)")
        btn_matrix.setProperty("btnType", "action")
        btn_matrix.setFixedHeight(34)
        btn_matrix.setMinimumWidth(330)
        btn_matrix.clicked.connect(lambda: self.open_matrix_requested.emit())

        top_bar.addWidget(lbl_title)
        top_bar.addStretch()
        top_bar.addWidget(btn_matrix)
        layout.addLayout(top_bar)

        # 2. Main Prompt Editor
        grp_main = QGroupBox("Main Prompt Editor / プロンプト編集")
        grp_layout = QVBoxLayout(grp_main)
        
        self.txt_main = PlainTextOnlyTextEdit()
        self.txt_main.setPlaceholderText("Enter prompts here, or click tags in the Dictionary Matrix...")
        grp_layout.addWidget(self.txt_main)

        layout.addWidget(grp_main, 3)

        # 3. DeepL Real-time Translation Assistant Panel (Split Layout)
        grp_trans = QGroupBox("🌐 DeepL Translation Assistant / リアルタイム翻訳アシスタント")
        grp_trans_layout = QVBoxLayout(grp_trans)
        grp_trans_layout.setContentsMargins(10, 8, 10, 8)
        grp_trans_layout.setSpacing(6)

        # Header bar of Translation panel
        trans_top = QHBoxLayout()
        lbl_hint = QLabel("<span style='color: #64748B; font-size: 11px;'>Translate words/phrases from your language into English prompt tags. / 母国語の表現を英訳してプロンプトや辞書に追加できます。</span>")
        
        btn_get_key = QPushButton("🔗 Get API Key (DeepL)")
        btn_get_key.setToolTip("Open DeepL API key page in browser (auto-redirects to your local language) / ブラウザでAPIキー取得ページを開く")
        btn_get_key.setFixedHeight(26)
        btn_get_key.clicked.connect(self.open_deepl_get_key_url)

        btn_set_key = QPushButton("⚙️ API Key")
        btn_set_key.setToolTip("Configure DeepL API Key / APIキーを設定・変更")
        btn_set_key.setFixedHeight(26)
        btn_set_key.clicked.connect(self.open_deepl_key_dialog)

        trans_top.addWidget(lbl_hint, 1)
        trans_top.addWidget(btn_get_key)
        trans_top.addWidget(btn_set_key)
        grp_trans_layout.addLayout(trans_top)

        # Source Language Input row
        h_ja = QHBoxLayout()
        lbl_ja = QLabel("<b>Source / 原文:</b>")
        lbl_ja.setFixedWidth(105)
        self.txt_trans_ja = QLineEdit()
        self.txt_trans_ja.setPlaceholderText("Enter word or phrase in your language (Enter to translate)... / 翻訳したい単語や表現")
        self.txt_trans_ja.returnPressed.connect(self.on_translate_exec)

        self.btn_trans_exec = QPushButton("🌐 Translate (翻訳)")
        self.btn_trans_exec.setProperty("btnType", "action")
        self.btn_trans_exec.setFixedHeight(30)
        self.btn_trans_exec.setMinimumWidth(130)
        self.btn_trans_exec.clicked.connect(self.on_translate_exec)

        h_ja.addWidget(lbl_ja)
        h_ja.addWidget(self.txt_trans_ja, 1)
        h_ja.addWidget(self.btn_trans_exec)
        grp_trans_layout.addLayout(h_ja)

        # English Output row & Quick Action Buttons
        h_en = QHBoxLayout()
        lbl_en = QLabel("<b>English / 英訳:</b>")
        lbl_en.setFixedWidth(105)
        self.txt_trans_en = QLineEdit()
        self.txt_trans_en.setPlaceholderText("English translation will appear here (Enter to add to prompt)... / 英訳結果")
        self.txt_trans_en.returnPressed.connect(self.on_add_translated_to_prompt)

        self.btn_add_trans = QPushButton("＋ Add to Prompt (追加)")
        self.btn_add_trans.setToolTip("Appends translated tag to Main Prompt with comma (respects active weight) / プロンプト末尾に追加")
        self.btn_add_trans.setProperty("btnType", "success")
        self.btn_add_trans.setFixedHeight(30)
        self.btn_add_trans.setMinimumWidth(150)
        self.btn_add_trans.clicked.connect(self.on_add_translated_to_prompt)

        self.btn_save_dict = QPushButton("💾 Save Tag (辞書登録)")
        self.btn_save_dict.setToolTip("Saves to User Dictionary (user_tags.json) and updates Matrix immediately / ユーザー辞書に保存")
        self.btn_save_dict.setFixedHeight(30)
        self.btn_save_dict.setMinimumWidth(130)
        self.btn_save_dict.clicked.connect(self.on_save_translated_to_dict)

        h_en.addWidget(lbl_en)
        h_en.addWidget(self.txt_trans_en, 1)
        h_en.addWidget(self.btn_add_trans)
        h_en.addWidget(self.btn_save_dict)
        grp_trans_layout.addLayout(h_en)

        layout.addWidget(grp_trans, 0)

        # 3. Weight & Modifier Controls
        weight_box = QHBoxLayout()
        
        self.chk_weight = QCheckBox("Enable Weight / 重み付け有効:")
        self.chk_weight.setChecked(False)

        self.cmb_weight = QComboBox()
        for w in ["0.5", "0.6", "0.7", "0.8", "0.9", "1.1", "1.2", "1.3", "1.4", "1.5"]:
            self.cmb_weight.addItem(w)
        self.cmb_weight.setCurrentText(self.config.get_setting("DefaultWeight", "1.1"))
        self.cmb_weight.setEnabled(False)

        self.btn_apply_weight = QPushButton("Apply Weight (tag:w)")
        self.btn_apply_weight.setEnabled(False)
        self.btn_apply_weight.clicked.connect(self.on_apply_weight)
        self.chk_weight.toggled.connect(self.btn_apply_weight.setEnabled)
        self.chk_weight.toggled.connect(self.cmb_weight.setEnabled)

        # Dynamic Prompts Wrap Button {A | B | C}
        btn_wrap = QPushButton("Wrap Dynamic Prompts {A | B | C}")
        btn_wrap.setToolTip("Converts selected comma-separated tags 'a, b, c' into '{a | b | c}' (or unwraps back)")
        btn_wrap.clicked.connect(self.on_wrap_dynamic_prompts)

        btn_sort = QPushButton("🔄 Sort Prompts (41 Categories)")
        btn_sort.setToolTip("Sorts tags strictly following Base Positive, LoRA triggers, and 41 Categories")
        btn_sort.setProperty("btnType", "action")
        btn_sort.clicked.connect(self.on_sort_prompt)

        weight_box.addWidget(self.chk_weight)
        weight_box.addWidget(self.cmb_weight)
        weight_box.addWidget(self.btn_apply_weight)
        weight_box.addWidget(btn_wrap)
        weight_box.addStretch()
        weight_box.addWidget(btn_sort)
        layout.addLayout(weight_box)

        # 4. Action Buttons Bar (Issue 1: Soft Clear with Warning Yellow & Hard Clear with Danger Red)
        action_bar = QHBoxLayout()

        self.btn_done = QPushButton("✨ Done! (Clean & Copy to Clipboard)")
        self.btn_done.setProperty("btnType", "success")
        self.btn_done.setFixedHeight(38)
        self.btn_done.clicked.connect(self.on_done)

        self.btn_add_fav = QPushButton("⭐ Send to Favorites / お気に入り転送")
        self.btn_add_fav.setFixedHeight(38)
        self.btn_add_fav.clicked.connect(self.on_add_to_fav)

        self.btn_undo = QPushButton("↩ Undo")
        self.btn_undo.setFixedHeight(38)
        self.btn_undo.clicked.connect(self.on_undo)

        # Soft Clear (Warning Yellow: preserves undo history)
        self.btn_clear_soft = QPushButton("🗑 Clear (履歴保持)")
        self.btn_clear_soft.setToolTip("Clears prompt editor while preserving undo history (Undoで復元可能)")
        self.btn_clear_soft.setFixedHeight(38)
        self.btn_clear_soft.setProperty("btnType", "warning")
        self.btn_clear_soft.clicked.connect(self.on_clear_soft)

        # Hard Clear (Danger Red: deletes undo history as well)
        self.btn_clear_hard = QPushButton("💣 Delete All (履歴ごと消去)")
        self.btn_clear_hard.setToolTip("Completely clears prompt editor AND all undo history (完全消去)")
        self.btn_clear_hard.setFixedHeight(38)
        self.btn_clear_hard.setProperty("btnType", "danger")
        self.btn_clear_hard.clicked.connect(self.on_clear_hard)

        action_bar.addWidget(self.btn_done, 2)
        action_bar.addWidget(self.btn_add_fav, 1)
        action_bar.addWidget(self.btn_undo)
        action_bar.addWidget(self.btn_clear_soft)
        action_bar.addWidget(self.btn_clear_hard)
        layout.addLayout(action_bar)

    def append_tag(self, tag_text: str, is_comma: bool = True):
        """Called when a tag is clicked from Matrix View or other tabs."""
        current = self.txt_main.toPlainText()
        self.engine.push_history(current)
        
        if self.chk_weight.isChecked():
            try:
                w_val = float(self.cmb_weight.currentText())
                tag_text = self.engine.format_weight(tag_text, w_val)
            except ValueError:
                pass
            self.chk_weight.setChecked(False)

        new_text = self.engine.append_tag(current, tag_text, is_comma=is_comma)
        self.txt_main.setPlainText(new_text)
        self.txt_main.moveCursor(QTextCursor.End)

    def insert_at_beginning(self, prefix_text: str):
        """Inserts text at the very beginning of the prompt."""
        current = self.txt_main.toPlainText().strip()
        prefix = prefix_text.strip()
        if not prefix:
            return
        self.engine.push_history(current)
        if current:
            combined = f"{prefix}, {current}"
        else:
            combined = prefix
    def receive_positive_prompt(self, pos_prompt: str) -> bool:
        """
        Inserts positive prompt at the beginning of Cockpit without asking overwrite/append.
        If EXACT SAME positive prompt is already in Cockpit, shows warning dialog and stops.
        Returns True if inserted successfully, False if aborted.
        """
        pos_clean = sanitize_sd_prompt(pos_prompt)
        if not pos_clean:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "ポジティブプロンプトが空です。\nPositive prompt is empty."
            )
            return False

        current = self.txt_main.toPlainText().strip()
        if not current:
            self.engine.push_history("")
            self.txt_main.setPlainText(pos_clean)
            return True

        # Duplicate check: check if exact pos_clean is equal to or present in current
        pos_norm = re.sub(r"\s+", " ", pos_clean.lower())
        curr_norm = re.sub(r"\s+", " ", current.lower())
        
        if pos_norm == curr_norm or pos_norm in curr_norm:
            QMessageBox.warning(
                self,
                "重複検知 / Duplicate Detected",
                "既に同じポジティブプロンプトがCockpit内に存在します。\n"
                "The exact same positive prompt is already present in Cockpit."
            )
            return False

        # Insert at the beginning with comma
        self.engine.push_history(current)
        combined = f"{pos_clean}, {current}"
        self.txt_main.setPlainText(combined)
        self.txt_main.moveCursor(QTextCursor.Start)
        return True

    def get_prompt(self) -> str:
        return self.txt_main.toPlainText().strip()

    def set_prompt(self, text: str):
        self.txt_main.setPlainText(sanitize_sd_prompt(text))

    def receive_prompt_from_external(self, new_prompt: str, is_prefix: bool = False):
        """
        Handles sending prompt from Favorites/Gacha with Overwrite / Append / Cancel dialog.
        """
        current = self.txt_main.toPlainText().strip()
        new_prompt = sanitize_sd_prompt(new_prompt)
        if not new_prompt:
            return

        if not current:
            self.txt_main.setPlainText(new_prompt)
            return

        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("送信確認 / Prompt Action")
        msg_box.setText(
            "コックピットに既にプロンプトが存在します。どのように処理しますか？\n"
            "The Cockpit prompt editor already contains text. How would you like to proceed?"
        )
        msg_box.setIcon(QMessageBox.Question)
        
        btn_overwrite = msg_box.addButton("上書き (Overwrite)", QMessageBox.AcceptRole)
        btn_append = msg_box.addButton("追記 (Append)", QMessageBox.ActionRole)
        btn_cancel = msg_box.addButton("キャンセル (Cancel)", QMessageBox.RejectRole)

        msg_box.exec()
        clicked = msg_box.clickedButton()

        if clicked == btn_overwrite:
            self.engine.push_history(current)
            self.txt_main.setPlainText(new_prompt)
        elif clicked == btn_append:
            self.engine.push_history(current)
            if is_prefix:
                self.txt_main.setPlainText(f"{new_prompt}, {current}")
            else:
                self.txt_main.setPlainText(self.engine.append_tag(current, new_prompt, is_comma=True))
        else:
            # Cancel
            pass

    def receive_lora_tag(self, lora_text: str):
        """
        Flexible smart LoRA transfer:
        Extracts each <lora:...> tag and trigger word independently.
        Skips only the already-present items in Cockpit, and appends NEW LoRA tags/triggers.
        If all items are already in Cockpit, notifies the user politely without failing.
        """
        current = self.txt_main.toPlainText().strip()
        lora_clean = lora_text.strip()
        if not lora_clean:
            return

        if not current:
            self.engine.push_history("")
            self.txt_main.setPlainText(lora_clean)
            self.txt_main.moveCursor(QTextCursor.End)
            return

        import re

        # Extract existing LoRA identifiers from current Cockpit prompt
        existing_lora_tags = re.findall(r"<lora:([^:>]+)(?::([0-9.]+))?>", current, re.IGNORECASE)
        existing_lora_names = {name.strip().lower() for name, _ in existing_lora_tags}

        # Parse transmission tokens into LoRA tags and trigger words
        # Regex captures <lora:...>, parenthesized weighted triggers (word:w), and standalone tokens
        tokens = [t.strip() for t in re.split(r"(<lora:[^>]+>|\([^)]+\))", lora_clean) if t.strip()]

        new_parts = []
        skipped_items = []

        for token in tokens:
            # Check if this token is a <lora:name:w> tag
            m_lora = re.match(r"^<lora:([^:>]+)(?::([0-9.]+))?>$", token, re.IGNORECASE)
            if m_lora:
                lora_name = m_lora.group(1).strip().lower()
                if lora_name in existing_lora_names:
                    skipped_items.append(token)
                else:
                    new_parts.append(token)
                    existing_lora_names.add(lora_name)
            else:
                # Sub-split triggers by comma/space if multiple words in one chunk
                sub_words = [w.strip() for w in re.split(r"[\s,]+", token) if w.strip()]
                for sw in sub_words:
                    # Clean weight brackets to check word presence e.g. (sound effects:1.2) -> sound effects
                    raw_term = re.sub(r"[():0-9.]", "", sw).strip().lower()
                    if raw_term and re.search(rf"(?:^|[,\s(]){re.escape(raw_term)}(?:[,\s):]|$)", current, re.IGNORECASE):
                        skipped_items.append(sw)
                    else:
                        new_parts.append(sw)

        if not new_parts:
            QMessageBox.information(
                self,
                "重複検知 / Already in Cockpit",
                f"送信されたLoRAタグは既にCockpit内にすべて含まれています:\nAll transmitted LoRA tags/triggers are already in Cockpit:\n\n{', '.join(skipped_items)}"
            )
            return

        # Append new parts
        self.engine.push_history(current)
        append_text = " ".join(new_parts)
        new_prompt = f"{current} {append_text}"
        self.txt_main.setPlainText(new_prompt)
        self.txt_main.moveCursor(QTextCursor.End)

        if skipped_items:
            QMessageBox.information(
                self,
                "追加完了（重複スキップ） / Added with Duplicates Skipped",
                f"新しいLoRAタグを追加しました！（重複タグはスキップ）:\nAdded new LoRA tags to Cockpit (skipped duplicates):\n\n"
                f"【追加】 {append_text}\n"
                f"【スキップ】 {', '.join(skipped_items)}"
            )

    def on_apply_weight(self):
        if not self.chk_weight.isChecked():
            return

        cursor = self.txt_main.textCursor()
        selected = cursor.selectedText()
        if not selected:
            QMessageBox.information(
                self,
                "情報 / Information",
                "重みを適用したいタグを選択してください。\nPlease select a tag/text in the prompt editor to apply weight."
            )
            return

        try:
            w_val = float(self.cmb_weight.currentText())
            weighted = self.engine.format_weight(selected, w_val)
            self.engine.push_history(self.txt_main.toPlainText())
            cursor.insertText(weighted)
            self.chk_weight.setChecked(False)
        except ValueError:
            pass

    def on_wrap_dynamic_prompts(self):
        cursor = self.txt_main.textCursor()
        selected = cursor.selectedText()
        if not selected:
            raw = self.txt_main.toPlainText().strip()
            if raw:
                self.engine.push_history(self.txt_main.toPlainText())
                wrapped = self.engine.wrap_dynamic_prompts(raw)
                self.txt_main.setPlainText(wrapped)
        else:
            self.engine.push_history(self.txt_main.toPlainText())
            wrapped = self.engine.wrap_dynamic_prompts(selected)
            cursor.insertText(wrapped)

    def on_sort_prompt(self):
        cursor = self.txt_main.textCursor()
        selected = cursor.selectedText().strip()
        lora_triggers = self.config.get_all_lora_triggers()

        if selected:
            # Sort ONLY the highlighted selection range
            # Note: Qt QTextCursor.selectedText uses \u2029 for paragraph separators
            norm_selected = selected.replace("\u2029", "\n")
            sorted_segment = self.engine.sort_prompt_segment(norm_selected, registered_lora_triggers=lora_triggers)
            self.engine.push_history(self.txt_main.toPlainText())
            cursor.insertText(sorted_segment)
            self.txt_main.setTextCursor(cursor)
        else:
            current = self.txt_main.toPlainText().strip()
            if not current:
                QMessageBox.warning(
                    self,
                    "注意 / Warning",
                    "プロンプトが空です。ソートする内容を入力してください。\nPrompt editor is empty. Nothing to sort."
                )
                return

            self.engine.push_history(self.txt_main.toPlainText())
            sorted_text = self.engine.sort_prompt(current, registered_lora_triggers=lora_triggers)
            self.txt_main.setPlainText(sorted_text)

    def on_done(self):
        current = self.txt_main.toPlainText()
        if not current.strip():
            return

        self.engine.push_history(current)
        cleaned = self.engine.clean_prompt(current)
        self.txt_main.setPlainText(cleaned)
        
        from .style import safe_copy_to_clipboard
        safe_copy_to_clipboard(cleaned, self)

        QMessageBox.information(
            self,
            "完了 / Success",
            "プロンプトを整形し、クリップボードにコピーしました！\nPrompt formatted and copied to clipboard!"
        )

    def on_undo(self):
        """Issue 4: Shows warning dialog when no undo history is available."""
        if not self.engine.history_stack:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "元に戻す履歴がありません。\nNo undo history available."
            )
            return

        prev = self.engine.undo()
        if prev is not None:
            self.txt_main.setPlainText(prev)
        else:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "元に戻す履歴がありません。\nNo undo history available."
            )

    def on_clear_soft(self):
        """Issue 4: Validates non-empty before soft clearing."""
        current = self.txt_main.toPlainText()
        if not current.strip():
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "プロンプトが空です。消去する内容がありません。\nPrompt editor is already empty."
            )
            return

        self.engine.push_history(current)
        self.txt_main.clear()

    def on_clear_hard(self):
        """Issue 4: Validates non-empty before hard clearing."""
        current = self.txt_main.toPlainText()
        if not current.strip():
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "プロンプトが空です。消去する内容がありません。\nPrompt editor is already empty."
            )
            return

        ans = QMessageBox.question(
            self,
            "⚠️ 履歴消去の確認 / Confirm Full Clear",
            "アンドゥ履歴を含めてプロンプトを完全に消去しますか？\nこの操作は取り消せません。\n\nAre you sure you want to clear prompt and ALL undo history?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if ans == QMessageBox.Yes:
            self.engine.clear_history()
            self.txt_main.clear()

    def on_add_to_fav(self):
        current = self.txt_main.toPlainText().strip()
        if not current:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "お気に入りに転送するプロンプトが空です。\nCannot send an empty prompt to Favorites."
            )
            return
        self.send_to_fav_requested.emit(current, "")

    # --- DeepL Translation & User Dictionary Operations ---
    def open_deepl_get_key_url(self):
        """Opens DeepL API key page in default browser (auto-redirects to local language)."""
        QDesktopServices.openUrl(QUrl(DeepLAPI.KEY_URL))

    def open_deepl_key_dialog(self):
        """Opens modal dialog to configure DeepL API key."""
        dlg = DeepLKeyDialog(self.config, self.deepl_api, self)
        dlg.exec()

    def on_translate_exec(self):
        """Starts asynchronous translation of Japanese input via DeepL API."""
        ja_text = self.txt_trans_ja.text().strip()
        if not ja_text:
            return

        api_key = self.config.get_setting("DeepLAPIKey", "").strip()
        if not api_key:
            # Helpful guidance dialog with direct links
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setWindowTitle("DeepL API Key Required / APIキー設定のお願い")
            msg_box.setText(
                "<b>DeepL API Key is missing. / APIキーが設定されていません。</b><br><br>"
                "To use the translation assistant, a DeepL API Free key (ends with <code>:fx</code>, 500,000 chars/month free) is required.<br>"
                "翻訳機能を利用するには、DeepL API Free（月50万文字無料）のAPIキーが必要です。<br><br>"
                "Would you like to get a key from DeepL official site now?<br>"
                "今すぐDeepL公式サイトでキーを発行しますか？"
            )
            btn_browser = msg_box.addButton("🔗 Get API Key (公式サイト)", QMessageBox.ActionRole)
            btn_enter_key = msg_box.addButton("⚙️ Enter Key (キー入力)", QMessageBox.ActionRole)
            msg_box.addButton("Cancel / キャンセル", QMessageBox.RejectRole)

            msg_box.exec()
            clicked_btn = msg_box.clickedButton()

            if clicked_btn == btn_browser:
                self.open_deepl_get_key_url()
                self.open_deepl_key_dialog()
            elif clicked_btn == btn_enter_key:
                self.open_deepl_key_dialog()
            return

        # Disable button while translating
        self.btn_trans_exec.setEnabled(False)
        self.btn_trans_exec.setText("Translating... / 翻訳中...")
        self.txt_trans_ja.setEnabled(False)

        self.deepl_api.set_api_key(api_key)
        self._trans_worker = DeepLTranslateWorker(self.deepl_api, ja_text, self)
        self._trans_worker.translation_finished.connect(self._on_translate_finished)
        self._trans_worker.translation_error.connect(self._on_translate_error)
        self._trans_worker.start()

    def _on_translate_finished(self, translated_text: str):
        self.btn_trans_exec.setEnabled(True)
        self.btn_trans_exec.setText("🌐 Translate (翻訳)")
        self.txt_trans_ja.setEnabled(True)
        self.txt_trans_en.setText(translated_text)
        self.txt_trans_en.setFocus()
        self.txt_trans_en.selectAll()

    def _on_translate_error(self, err_msg: str):
        self.btn_trans_exec.setEnabled(True)
        self.btn_trans_exec.setText("🌐 Translate (翻訳)")
        self.txt_trans_ja.setEnabled(True)
        QMessageBox.warning(
            self,
            "翻訳エラー / Translation Error",
            f"翻訳に失敗しました:\nTranslation failed:\n\n{err_msg}"
        )

    def on_add_translated_to_prompt(self):
        """Adds translated English text to Main Prompt Editor."""
        en_text = self.txt_trans_en.text().strip()
        if not en_text:
            return

        self.append_tag(en_text, is_comma=True)
        
        # Clear translation inputs for next query
        self.txt_trans_ja.clear()
        self.txt_trans_en.clear()
        self.txt_trans_ja.setFocus()

    def on_save_translated_to_dict(self):
        """Saves current translation pair to User Dictionary (user_tags.json) and notifies Matrix."""
        ja_text = self.txt_trans_ja.text().strip()
        en_text = self.txt_trans_en.text().strip()

        if not en_text:
            QMessageBox.warning(
                self,
                "注意 / Warning",
                "辞書に登録する英語プロンプトが空です。\n英訳を実行するか、直接入力してください。\n\n"
                "English prompt is empty. Please translate or enter a prompt."
            )
            return

        if not self.db:
            QMessageBox.critical(self, "エラー / Error", "データベースマネージャーが初期化されていません。(DatabaseManager is not initialized)")
            return

        success, new_tag, msg = self.db.add_user_tag(label_ja=ja_text, prompt_en=en_text, note="DeepL Translation")
        if success and new_tag:
            self.user_tag_added.emit(new_tag)
            QMessageBox.information(
                self,
                "辞書登録完了 / Registered to Dictionary",
                f"ユーザー辞書（user_tags.json）に登録しました！\n"
                f"辞書マトリクス画面および検索窓ですぐに利用できます。\n\n"
                f"Saved to User Dictionary! Available immediately in Matrix and search.\n\n"
                f"【Source / 原文】: {ja_text or '(None / 未指定)'}\n"
                f"【English / 英訳】: {en_text}"
            )
        else:
            QMessageBox.warning(self, "登録失敗 / Failed", msg)
