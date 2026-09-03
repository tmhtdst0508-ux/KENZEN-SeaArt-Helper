"""
Dictionary Matrix Window for KENZEN SeaArt Helper v5.0.0
Full-width scrollable tag matrix interface with category-specific header colors,
compact hints, and Category Jump combo box.
"""

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QScrollArea, QFrame, QGridLayout, QComboBox, QTabWidget, QSplitter,
    QTableWidget, QTableWidgetItem, QHeaderView, QTextEdit, QApplication, QMessageBox
)
from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QIcon, QMouseEvent
from ..core.db_manager import DBManager
from .style import (
    COLOR_ACTION, COLOR_TAG_EN_BG, COLOR_TAG_EN_HOVER, COLOR_TAG_JP_BG
)
from .widgets import PlainTextOnlyTextEdit


def get_category_header_color(order: int) -> str:
    """Returns the specific category header color designated in issue 1-3-2."""
    if order == 0:
        return "#C1F0C8" # 0. 接続助詞・前置詞
    elif 1 <= order <= 2:
        return "#BAE6FD" # 1. 画風 ～ 2. アングル
    elif 3 <= order <= 7:
        return "#FEF08A" # 3. キャラ数 ～ 7. 体型
    elif 8 <= order <= 13:
        return "#A7F3D0" # 8. 髪の毛 ～ 13. 瞳の色
    elif 14 <= order <= 20:
        return "#BBF7D0" # 14. 表情 ～ 20. ボンデージ行為
    elif 21 <= order <= 30:
        return "#FECDD3" # 21. 職業 ～ 30. アクセサリー
    elif 31 <= order <= 34:
        return "#6EE7B7" # 31. 手段・道具 ～ 34. 体液
    elif 35 <= order <= 36:
        return "#E9D5FF" # 35. 場所 ～ 36. 時間帯・周囲の状況
    elif order == 37:
        return "#FFEDD5" # 37. その他アイテム
    else:
        return "#CBD5E1" # 38. 光源 ～ 40. 修正/特殊


class MatrixTagButton(QPushButton):
    """Custom button handling Left-Click (Comma) and Right-Click (Space)."""
    clicked_custom = Signal(str, bool) # (tag_text, is_comma)

    def __init__(self, tag_text: str, parent=None):
        super().__init__(tag_text, parent)
        self.tag_text = tag_text

    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            self.clicked_custom.emit(self.tag_text, True)
        elif event.button() == Qt.RightButton:
            self.clicked_custom.emit(self.tag_text, False)
        super().mousePressEvent(event)


class MatrixWindow(QWidget):
    tag_selected = Signal(str, bool)           # (tag_text, is_comma)
    sample_prompt_selected = Signal(str, str)  # (prompt_text, target_tab) e.g. target_tab="cockpit" or "positive"

    def __init__(self, db_manager: DBManager, parent=None):
        super().__init__(parent)
        self.db = db_manager
        self.category_widgets: dict = {}
        self.tag_widget_map: dict = {} # tag_id -> (tag_box, btn_en, tag_dict)
        self.search_results: list = []
        self.search_index: int = -1
        self.last_search_kw: str = ""
        self.flash_step: int = 0
        self.current_flashing_widgets: tuple = None
        self.flash_timer = QTimer(self)
        self.flash_timer.setInterval(160)
        self.flash_timer.timeout.connect(self._on_flash_tick)
        
        self.sample_prompts_data: list = []
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("KENZEN Tag & Sample Database / 辞書マトリクス ＆ サンプルプロンプト")
        self.resize(1100, 780)
        self.setMinimumSize(750, 520)

        # Set Window Icon
        from ..core.path_utils import get_resource_path
        icon_path = get_resource_path("kenzen_icon.ico")
        import os
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(6, 6, 6, 6)
        main_layout.setSpacing(6)

        # Main Tab Widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabBar::tab {
                padding: 7px 18px;
                font-weight: bold;
                font-size: 13px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                background-color: #E2E8F0;
                color: #475569;
                margin-right: 3px;
            }
            QTabBar::tab:selected {
                background-color: #FFFFFF;
                color: #1E3A8A;
                border-bottom: 2px solid #2563EB;
            }
        """)

        # Tab 1: Tag Matrix
        tab_matrix = QWidget()
        self._init_tag_matrix_tab(tab_matrix)
        self.tab_widget.addTab(tab_matrix, "🏷️ タグ辞書マトリクス (Tag Matrix)")

        # Tab 2: Sample Prompts
        tab_samples = QWidget()
        self._init_sample_prompts_tab(tab_samples)
        self.tab_widget.addTab(tab_samples, "📖 サンプルプロンプト集 (Sample Prompts)")

        main_layout.addWidget(self.tab_widget)

    def _init_tag_matrix_tab(self, parent_widget: QWidget):
        layout = QVBoxLayout(parent_widget)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # 1. Top Control Bar (Search, Next Search, Jump, and Compact Hint)
        top_bar = QHBoxLayout()
        top_bar.setSpacing(6)
        
        lbl_search = QLabel("🔍 Search:")
        lbl_search.setStyleSheet("font-weight: bold; font-size: 12px;")
        
        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Search tag / note (Press Enter)...")
        self.txt_search.setMinimumWidth(130)
        self.txt_search.returnPressed.connect(self.on_search_execute)
        self.txt_search.textChanged.connect(self.on_search_text_changed)
        
        btn_search = QPushButton("Next ▶")
        btn_search.setProperty("btnType", "action")
        btn_search.setToolTip("Search or jump to next match (Enter)")
        btn_search.clicked.connect(self.on_search_execute)

        btn_clear_search = QPushButton("Clear")
        btn_clear_search.setProperty("btnType", "danger")
        btn_clear_search.clicked.connect(self.on_search_clear)

        lbl_jump = QLabel("Jump:")
        lbl_jump.setStyleSheet("font-weight: bold; font-size: 12px; margin-left: 4px;")
        
        self.cmb_categories = QComboBox()
        self.cmb_categories.setMinimumWidth(240)
        self.cmb_categories.currentIndexChanged.connect(self.on_category_jump)

        # Compact Hint badge
        lbl_hint = QLabel("💡 <b>Left:</b> (, ) | <b>Right:</b> ( )")
        lbl_hint.setStyleSheet("color: #1E3A8A; font-size: 11px; background-color: #EFF6FF; border: 1px solid #BFDBFE; padding: 3px 6px; border-radius: 3px;")

        top_bar.addWidget(lbl_search)
        top_bar.addWidget(self.txt_search, 1)
        top_bar.addWidget(btn_search)
        top_bar.addWidget(btn_clear_search)
        top_bar.addWidget(lbl_jump)
        top_bar.addWidget(self.cmb_categories)
        top_bar.addWidget(lbl_hint)

        layout.addLayout(top_bar)

        # 2. Full-Width Scroll Area for Matrix Grid
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scroll_content = QWidget()
        self.matrix_layout = QVBoxLayout(self.scroll_content)
        self.matrix_layout.setSpacing(14)
        self.matrix_layout.setContentsMargins(6, 6, 6, 6)
        
        self.scroll_area.setWidget(self.scroll_content)
        layout.addWidget(self.scroll_area, 1)

        self.populate_matrix()

    def _init_sample_prompts_tab(self, parent_widget: QWidget):
        layout = QVBoxLayout(parent_widget)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)

        # 1. Top Control Bar (Search & Reload)
        top_bar = QHBoxLayout()
        top_bar.setSpacing(8)

        lbl_filter = QLabel("🔍 Filter / 検索:")
        lbl_filter.setStyleSheet("font-weight: bold; font-size: 12px;")
        top_bar.addWidget(lbl_filter)

        self.txt_sample_search = QLineEdit()
        self.txt_sample_search.setPlaceholderText("Filter sample by keyword / タイトル・説明、プロンプト、作者コメントで絞り込み...")
        self.txt_sample_search.textChanged.connect(self._filter_sample_prompts)
        top_bar.addWidget(self.txt_sample_search, 1)

        btn_reload = QPushButton("🔄 Reload / 再読込")
        btn_reload.setProperty("btnType", "action")
        btn_reload.clicked.connect(self.load_sample_prompts)
        top_bar.addWidget(btn_reload)

        layout.addLayout(top_bar)

        # 2. Main Content Splitter (Left: Table, Right: Detail View & Actions)
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        # Left Panel: Table
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        lbl_table_header = QLabel("📋 Sample Prompts List / サンプル一覧 (Click to view, Double-click to send)")
        lbl_table_header.setStyleSheet("color: #475569; font-size: 11px; font-weight: 600;")
        left_layout.addWidget(lbl_table_header)

        self.tbl_samples = QTableWidget()
        self.tbl_samples.setColumnCount(3)
        self.tbl_samples.setHorizontalHeaderLabels(["ID", "Title & Description / サンプル説明", "Author Comment / 作者コメント"])
        self.tbl_samples.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_samples.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tbl_samples.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.tbl_samples.setColumnWidth(2, 220)
        self.tbl_samples.verticalHeader().setVisible(False)
        self.tbl_samples.setSelectionBehavior(QTableWidget.SelectRows)
        self.tbl_samples.setSelectionMode(QTableWidget.SingleSelection)
        self.tbl_samples.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_samples.setAlternatingRowColors(True)
        self.tbl_samples.itemSelectionChanged.connect(self._on_sample_selected)
        self.tbl_samples.itemDoubleClicked.connect(lambda item: self._on_send_to_cockpit())
        left_layout.addWidget(self.tbl_samples, 1)

        splitter.addWidget(left_panel)

        # Right Panel: Detail Preview & Actions
        right_panel = QFrame()
        right_panel.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 6px;
            }
        """)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(8, 8, 8, 8)
        right_layout.setSpacing(6)

        lbl_preview_title = QLabel("📖 Selected Prompt Details / 選択プロンプト詳細")
        lbl_preview_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #1E3A8A;")
        right_layout.addWidget(lbl_preview_title)

        # Title & Description
        self.lbl_sample_desc = QLabel("タイトル・説明がここに表示されます")
        self.lbl_sample_desc.setStyleSheet("font-size: 12px; color: #1E293B; font-weight: 600; padding: 6px; background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 4px;")
        self.lbl_sample_desc.setWordWrap(True)
        right_layout.addWidget(self.lbl_sample_desc)

        # Author Comment Box
        self.lbl_sample_comment = QLabel("💡 作者コメントがここに表示されます")
        self.lbl_sample_comment.setStyleSheet("font-size: 11px; color: #166534; background-color: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 4px; padding: 5px; font-weight: 500;")
        self.lbl_sample_comment.setWordWrap(True)
        right_layout.addWidget(self.lbl_sample_comment)

        lbl_prompt_header = QLabel("Prompt Content / プロンプト本文:")
        lbl_prompt_header.setStyleSheet("font-size: 11px; font-weight: bold; color: #475569; margin-top: 2px;")
        right_layout.addWidget(lbl_prompt_header)

        self.txt_sample_prompt = PlainTextOnlyTextEdit()
        self.txt_sample_prompt.setPlaceholderText("Select a sample prompt from the table on the left...")
        self.txt_sample_prompt.setStyleSheet("font-size: 12px; line-height: 1.4; background-color: #FFFFFF;")
        right_layout.addWidget(self.txt_sample_prompt, 1)

        # Action Buttons
        btn_bar = QHBoxLayout()
        btn_bar.setSpacing(8)

        btn_copy = QPushButton("📋 Copy / コピー")
        btn_copy.setProperty("btnType", "default")
        btn_copy.clicked.connect(self._on_copy_sample_prompt)
        btn_bar.addWidget(btn_copy)

        btn_send_cockpit = QPushButton("🚀 Cockpitへ転送")
        btn_send_cockpit.setProperty("btnType", "primary")
        btn_send_cockpit.clicked.connect(self._on_send_to_cockpit)
        btn_bar.addWidget(btn_send_cockpit)

        right_layout.addLayout(btn_bar)

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        layout.addWidget(splitter, 1)

        # Load samples
        self.load_sample_prompts()

    def load_sample_prompts(self):
        """Loads sample prompts from DBManager (SQLite sample_prompts table or JSON fallback)."""
        self.sample_prompts_data = self.db.get_sample_prompts()
        self._filter_sample_prompts()

    def _filter_sample_prompts(self):
        """Filters sample prompts by search keyword across title, prompt, and comment."""
        kw = self.txt_sample_search.text().strip().lower()

        filtered = []
        for s in self.sample_prompts_data:
            title = str(s.get("title", ""))
            prompt = str(s.get("prompt", ""))
            comment = str(s.get("comment", ""))

            if kw:
                if kw not in title.lower() and kw not in prompt.lower() and kw not in comment.lower():
                    continue

            filtered.append(s)

        self.tbl_samples.setRowCount(len(filtered))
        for row_idx, s in enumerate(filtered):
            item_id = QTableWidgetItem(str(s.get("id", row_idx + 1)))
            item_id.setTextAlignment(Qt.AlignCenter)
            item_title = QTableWidgetItem(str(s.get("title", "")))
            item_comment = QTableWidgetItem(str(s.get("comment", "")))
            
            # Store full data in item
            item_id.setData(Qt.UserRole, s)

            self.tbl_samples.setItem(row_idx, 0, item_id)
            self.tbl_samples.setItem(row_idx, 1, item_title)
            self.tbl_samples.setItem(row_idx, 2, item_comment)

        if filtered:
            self.tbl_samples.selectRow(0)
        else:
            self.lbl_sample_desc.setText("一致するサンプルプロンプトがありません")
            self.lbl_sample_comment.setText("💡 作者コメントなし")
            self.txt_sample_prompt.clear()

    def _on_sample_selected(self):
        """Updates preview when user selects a row in sample table."""
        selected_items = self.tbl_samples.selectedItems()
        if not selected_items:
            return

        row = selected_items[0].row()
        id_item = self.tbl_samples.item(row, 0)
        if not id_item:
            return

        data = id_item.data(Qt.UserRole)
        if data:
            self.lbl_sample_desc.setText(str(data.get("title", "")))
            comment = str(data.get("comment", "")).strip()
            if comment:
                self.lbl_sample_comment.setText(f"💡 <b>作者コメント:</b> {comment}")
                self.lbl_sample_comment.setVisible(True)
            else:
                self.lbl_sample_comment.setText("💡 作者コメントなし")
                self.lbl_sample_comment.setVisible(True)

            self.txt_sample_prompt.setPlainText(str(data.get("prompt", "")))

    def _on_copy_sample_prompt(self):
        text = self.txt_sample_prompt.toPlainText().strip()
        if text:
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
            QMessageBox.information(self, "コピー完了 / Copied", "サンプルプロンプトをクリップボードにコピーしました！")

    def _on_send_to_cockpit(self):
        text = self.txt_sample_prompt.toPlainText().strip()
        if text:
            self.sample_prompt_selected.emit(text, "cockpit")

    def populate_matrix(self):
        """Populates all 41 categories and their tags with custom color headers."""
        categories = self.db.get_categories()
        
        self.cmb_categories.blockSignals(True)
        self.cmb_categories.clear()
        self.cmb_categories.addItem("-- Jump to Category / カテゴリ選択 --", None)

        # Clear existing layout
        while self.matrix_layout.count():
            item = self.matrix_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.category_widgets.clear()

        for cat in categories:
            cid = cat["id"]
            order = cat["category_order"]
            cname = cat["category_name"]
            tags = self.db.get_tags_by_category(cid)
            
            display_title = f"{order}. {cname} ({len(tags)})"
            self.cmb_categories.addItem(display_title, cid)

            # Category Card Frame
            card = QFrame()
            card.setFrameShape(QFrame.StyledPanel)
            card.setStyleSheet("""
                QFrame {
                    background-color: #FFFFFF;
                    border: 1px solid #CBD5E1;
                    border-radius: 6px;
                }
            """)
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(6, 6, 6, 6)
            card_layout.setSpacing(6)

            # Header with specific color from Issue 1-3-2
            header_color = get_category_header_color(order)
            header = QLabel(f"  {display_title}")
            header.setStyleSheet(f"""
                background-color: {header_color};
                color: #0F172A;
                font-weight: bold;
                font-size: 13px;
                padding: 6px 8px;
                border-radius: 4px;
                border: 1px solid #94A3B8;
            """)
            card_layout.addWidget(header)

            # Tags Grid (4 columns for maximum screen efficiency)
            grid_widget = QWidget()
            grid = QGridLayout(grid_widget)
            grid.setContentsMargins(2, 2, 2, 2)
            grid.setSpacing(6)

            col_count = 4
            for c in range(col_count):
                grid.setColumnStretch(c, 1)

            for idx, tag in enumerate(tags):
                row = idx // col_count
                col = idx % col_count

                tag_box = QFrame()
                tag_box.setStyleSheet(f"""
                    QFrame {{
                        background-color: {COLOR_TAG_JP_BG};
                        border: 1px solid #E2E8F0;
                        border-radius: 4px;
                    }}
                """)
                t_layout = QVBoxLayout(tag_box)
                t_layout.setContentsMargins(4, 4, 4, 4)
                t_layout.setSpacing(2)

                # Japanese label
                ja_label_text = tag.get("label_ja", "")
                note_text = str(tag.get("note") or "").strip()

                if note_text:
                    lbl_ja = QLabel(f"{ja_label_text} 💡")
                else:
                    lbl_ja = QLabel(ja_label_text)

                lbl_ja.setStyleSheet("font-size: 11px; color: #475569; font-weight: 500;")
                lbl_ja.setWordWrap(True)
                t_layout.addWidget(lbl_ja)

                # English prompt button
                en_text = tag.get("prompt_en", "")
                btn_en = MatrixTagButton(en_text)

                if note_text:
                    tooltip_str = f"💡 [解説 / Note]\n{note_text}\n\n🖱️ Left-Click: Add with Comma (, )\n🖱️ Right-Click: Add with Space ( )\n[{en_text}]"
                    btn_en.setToolTip(tooltip_str)
                    tag_box.setToolTip(tooltip_str)
                    lbl_ja.setToolTip(tooltip_str)
                else:
                    btn_en.setToolTip(f"🖱️ Left-Click: Add with Comma (, )\n🖱️ Right-Click: Add with Space ( )\n[{en_text}]")

                btn_en.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {COLOR_TAG_EN_BG};
                        border: 1px solid #93C5FD;
                        border-radius: 3px;
                        padding: 4px 6px;
                        font-weight: 600;
                        font-size: 12px;
                        color: #0F172A;
                        text-align: left;
                    }}
                    QPushButton:hover {{
                        background-color: {COLOR_TAG_EN_HOVER};
                    }}
                """)
                btn_en.clicked_custom.connect(
                    lambda tag_text, is_comma, tb=tag_box, b=btn_en: self.on_tag_button_clicked(tag_text, is_comma, tb, b)
                )
                t_layout.addWidget(btn_en)

                grid.addWidget(tag_box, row, col)
                self.tag_widget_map[tag["id"]] = (cid, tag_box, btn_en, tag)

            card_layout.addWidget(grid_widget)
            self.matrix_layout.addWidget(card)
            self.category_widgets[cid] = card

        self.matrix_layout.addStretch()
        self.cmb_categories.blockSignals(False)

    def on_tag_button_clicked(self, tag_text: str, is_comma: bool, tag_box: QFrame = None, btn_en: MatrixTagButton = None):
        self.tag_selected.emit(tag_text, is_comma)
        if tag_box and btn_en:
            self.flash_tag_widget(tag_box, btn_en)

    def scroll_to_category(self, cid: int):
        """Scrolls directly to top-left of category card."""
        if cid is None or cid not in self.category_widgets:
            if self.cmb_categories.currentIndex() == 0:
                self.scroll_area.verticalScrollBar().setValue(0)
                self.scroll_area.horizontalScrollBar().setValue(0)
            return

        widget = self.category_widgets[cid]
        top_left = widget.mapTo(self.scroll_content, widget.rect().topLeft())
        
        if top_left.y() <= 20 or self.cmb_categories.currentIndex() == 0:
            target_y = 0
        else:
            target_y = max(0, top_left.y() - 4)

        self.scroll_area.verticalScrollBar().setValue(target_y)
        self.scroll_area.horizontalScrollBar().setValue(0)

    def scroll_to_tag_widget(self, widget: QWidget):
        """Scrolls smoothly and reliably so that the target tag widget is clearly visible."""
        if not widget:
            return
        QApplication.processEvents()
        self.scroll_area.ensureWidgetVisible(widget, 50, 80)

    def on_category_jump(self, index: int):
        cid = self.cmb_categories.currentData()
        if cid is not None:
            self.scroll_to_category(cid)
        elif index == 0:
            self.scroll_area.verticalScrollBar().setValue(0)
            self.scroll_area.horizontalScrollBar().setValue(0)

    def on_search_text_changed(self, text: str):
        """Resets search state when keyword changes."""
        kw = text.strip()
        if kw != self.last_search_kw:
            self.search_results = []
            self.search_index = -1

    def on_search_clear(self):
        """Clears search input, resets results, and resets flash animation."""
        self.txt_search.clear()
        self.search_results = []
        self.search_index = -1
        self.last_search_kw = ""
        self.flash_timer.stop()
        if self.current_flashing_widgets:
            prev_box, prev_btn = self.current_flashing_widgets
            self._restore_widget_style(prev_box, prev_btn)
            self.current_flashing_widgets = None

    def on_search_execute(self):
        """
        Executes tag search:
        - On first search with new keyword: searches all tags (ja, en, note), highlights and scrolls to match #1.
        - On subsequent Enter / Next click: sequentially loops to match #2, #3, ... #1 with visible yellow flash.
        """
        kw = self.txt_search.text().strip()
        if not kw:
            return

        if kw != self.last_search_kw or not self.search_results:
            self.last_search_kw = kw
            matched_tags = self.db.search_tags(kw)
            self.search_results = []
            for t in matched_tags:
                tid = t.get("id")
                if tid in self.tag_widget_map:
                    self.search_results.append(self.tag_widget_map[tid])
            self.search_index = 0
        else:
            # Sequential Loop: next item in results
            if len(self.search_results) > 0:
                self.search_index = (self.search_index + 1) % len(self.search_results)

        if not self.search_results:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.information(
                self,
                "検索結果 / Search",
                f"'{kw}' に一致するタグは見つかりませんでした。\nNo matching tags found."
            )
            return

        # Get matched item
        cid, tag_box, btn_en, tag_dict = self.search_results[self.search_index]
        
        # Scroll to matched widget
        self.scroll_to_tag_widget(tag_box)

        # Trigger yellow flash animation (3 pulses for search)
        self.flash_tag_widget(tag_box, btn_en, pulses=3)

    def flash_tag_widget(self, tag_box: QFrame, btn_en: MatrixTagButton, pulses: int = 1):
        """Starts yellow flashing animation on target tag cell (pulses=1 for click, pulses=3 for search)."""
        self.flash_timer.stop()
        # Reset previous flashing widget if any
        if self.current_flashing_widgets:
            prev_box, prev_btn = self.current_flashing_widgets
            self._restore_widget_style(prev_box, prev_btn)

        self.current_flashing_widgets = (tag_box, btn_en)
        self.flash_step = 0
        self.flash_max_steps = pulses * 2
        self.flash_timer.setInterval(160 if pulses > 1 else 220)
        self.flash_timer.start()
        # Immediate first tick for instant visual feedback
        self._on_flash_tick()

    def _on_flash_tick(self):
        """Timer callback to toggle yellow flash highlighting."""
        if not self.current_flashing_widgets:
            self.flash_timer.stop()
            return

        tag_box, btn_en = self.current_flashing_widgets
        self.flash_step += 1

        if self.flash_step % 2 == 1:
            # Bright Yellow Flash Highlight
            tag_box.setStyleSheet("""
                QFrame {
                    background-color: #FEF08A;
                    border: 2px solid #F59E0B;
                    border-radius: 4px;
                }
            """)
            btn_en.setStyleSheet("""
                QPushButton {
                    background-color: #FDE047;
                    border: 1.5px solid #D97706;
                    border-radius: 3px;
                    padding: 4px 6px;
                    font-weight: bold;
                    font-size: 12px;
                    color: #78350F;
                    text-align: left;
                }
            """)
        else:
            self._restore_widget_style(tag_box, btn_en)

        if self.flash_step >= getattr(self, "flash_max_steps", 2):
            self.flash_timer.stop()
            self._restore_widget_style(tag_box, btn_en)
            self.current_flashing_widgets = None

    def _restore_widget_style(self, tag_box: QFrame, btn_en: MatrixTagButton):
        """Restores normal clean styles to tag box and button."""
        tag_box.setStyleSheet(f"""
            QFrame {{
                background-color: {COLOR_TAG_JP_BG};
                border: 1px solid #E2E8F0;
                border-radius: 4px;
            }}
        """)
        btn_en.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLOR_TAG_EN_BG};
                border: 1px solid #93C5FD;
                border-radius: 3px;
                padding: 4px 6px;
                font-weight: 600;
                font-size: 12px;
                color: #0F172A;
                text-align: left;
            }}
            QPushButton:hover {{
                background-color: {COLOR_TAG_EN_HOVER};
            }}
        """)
