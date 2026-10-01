"""
Database Manager for KENZEN SeaArt Helper v5.0.0
Handles SQLite tags.db operations, category loading, tag retrieval, and search.
"""

import os
import sys
import json
import sqlite3
import datetime
from typing import List, Dict, Tuple, Optional, Any


from .path_utils import get_resource_path

USER_CATEGORY_ID = 9999
USER_CATEGORY_NAME = "★ ユーザー登録タグ (User Dictionary)"
USER_CATEGORY_ORDER = 41


class DBManager:
    def __init__(self, db_path: Optional[str] = None):
        if db_path is None:
            self.db_path = get_resource_path("tags.db")
        else:
            self.db_path = db_path

        self.user_tags_path = get_resource_path("user_tags.json")
        self._user_tags_cache: List[Dict] = []

        self._categories_cache: List[Dict] = []
        self._tags_by_category_cache: Dict[int, List[Dict]] = {}
        self._jp_to_en_map: Dict[str, str] = {}
        self._en_to_jp_map: Dict[str, str] = {}
        self._tag_to_category_order: Dict[str, int] = {}
        self._tag_to_category_order_collapsed: Dict[str, int] = {}
        
        # Verify database health before loading
        is_healthy, _ = self.verify_database_health()
        if is_healthy:
            self.load_all_data()

    def verify_database_health(self) -> Tuple[bool, str]:
        """
        Performs 3-tier integrity checks on SQLite database:
        1. File existence and size (> 1KB)
        2. SQLite PRAGMA integrity_check == 'ok'
        3. Essential tables ('categories', 'tags') and non-zero rows exist.
        Returns (is_healthy: bool, detail_message: str).
        """
        # 1. Existence and size check (prevents sqlite3.connect from creating a 0-byte fake DB)
        if not os.path.exists(self.db_path):
            return False, "データベースファイルが見つかりません (File not found)"
        
        try:
            if os.path.getsize(self.db_path) < 1024:
                return False, "データベースファイルが空、またはサイズが小さすぎます (Empty or incomplete file)"
        except Exception as e:
            return False, f"ファイルアクセスエラー (File access error: {e})"

        # 2. SQLite integrity check
        conn = None
        try:
            conn = sqlite3.connect(self.db_path, timeout=5.0)
            cur = conn.cursor()
            cur.execute("PRAGMA integrity_check;")
            res = cur.fetchone()
            if not res or res[0] != "ok":
                return False, f"データベース内部の破損を検知しました (Integrity check failed: {res[0] if res else 'Unknown'})"

            # 3. Essential tables & records check
            cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = [r[0] for r in cur.fetchall()]
            if "categories" not in tables or "tags" not in tables:
                return False, "必要なテーブル（categories, tags）が見つかりません (Missing required tables)"

            cur.execute("SELECT count(*) FROM tags;")
            tag_count = cur.fetchone()[0]
            if tag_count == 0:
                return False, "タグ辞書データが0件です (No tag records found in database)"

        except Exception as e:
            return False, f"データベース接続・検証エラー (Database error: {e})"
        finally:
            if conn:
                try:
                    conn.close()
                except Exception:
                    pass

        return True, "OK"

    def enforce_valid_database_or_exit(self):
        """Checks database health. If invalid, shows a critical dialog and terminates the process immediately."""
        is_healthy, detail = self.verify_database_health()
        if is_healthy:
            return

        import sys
        from PySide6.QtWidgets import QApplication, QMessageBox

        app = QApplication.instance() or QApplication(sys.argv)
        msg = (
            "【⚠️ 辞書データベースの異常 / Database Error】\n\n"
            "タグ辞書データベース（tags.db）が見つからないか、破損しています。\n"
            "The tag dictionary database (tags.db) was not found or is corrupted.\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"【詳細 / Detail】\n{detail}\n\n"
            "【対処方法 / Solution】\n"
            "本ツールのフォルダ（EXEと同じ場所）に、正常な 'tags.db' を\n"
            "配置してから、再度起動してください。\n"
            "Please place a valid 'tags.db' file in the application folder.\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            f"Path: {os.path.abspath(self.db_path)}"
        )

        box = QMessageBox()
        box.setIcon(QMessageBox.Critical)
        box.setWindowTitle("データベースエラー / Database Error")
        box.setText(msg)
        box.setStandardButtons(QMessageBox.Ok)
        box.exec()

        sys.exit(1)

    def get_connection(self) -> sqlite3.Connection:
        # Edge-case Guard 1: Timeout and WAL/busy_timeout for concurrent DB access (e.g. DB Browser)
        conn = sqlite3.connect(self.db_path, timeout=10.0)
        conn.execute("PRAGMA busy_timeout = 5000;")
        conn.row_factory = sqlite3.Row
        return conn

    def load_all_data(self):
        """Loads and caches all categories, tags, and mapping dictionaries."""
        if not os.path.exists(self.db_path):
            print(f"[DBManager] Warning: Database file not found at {self.db_path}")
            return

        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                # Load Categories
                cursor.execute("SELECT id, category_order, category_name FROM categories ORDER BY category_order ASC")
                cats = cursor.fetchall()
                self._categories_cache = [dict(row) for row in cats]

                # Check if 'note' column exists
                cursor.execute("PRAGMA table_info(tags)")
                cols = [c["name"] for c in cursor.fetchall()]
                has_note = "note" in cols

                if has_note:
                    cursor.execute("SELECT id, category_id, category_name, label_ja, prompt_en, note FROM tags ORDER BY prompt_en COLLATE NOCASE ASC")
                else:
                    cursor.execute("SELECT id, category_id, category_name, label_ja, prompt_en FROM tags ORDER BY prompt_en COLLATE NOCASE ASC")
                
                tags = cursor.fetchall()
                
                self._tags_by_category_cache.clear()
                self._jp_to_en_map.clear()
                self._en_to_jp_map.clear()
                self._tag_to_category_order.clear()
                self._tag_to_category_order_collapsed.clear()

                cat_order_map = {c["id"]: c["category_order"] for c in self._categories_cache}

                for t in tags:
                    tag_dict = dict(t)
                    if "note" not in tag_dict or tag_dict["note"] is None:
                        tag_dict["note"] = ""

                    cid = tag_dict["category_id"]
                    if cid not in self._tags_by_category_cache:
                        self._tags_by_category_cache[cid] = []
                    self._tags_by_category_cache[cid].append(tag_dict)

                    ja = tag_dict.get("label_ja", "").strip()
                    en = tag_dict.get("prompt_en", "").strip()

                    if ja and en:
                        self._jp_to_en_map[ja] = en
                        self._en_to_jp_map[en.lower()] = ja

                    # Record category order for sorting prompt tags
                    order = cat_order_map.get(cid, 999)
                    if order == 0:
                        order = 990 # Category 0 (particles/prepositions) sorted toward the end, not top
                    if en:
                        for token in en.split(","):
                            clean_tok = token.strip().lower()
                            if clean_tok and clean_tok not in self._tag_to_category_order:
                                self._tag_to_category_order[clean_tok] = order
                            col_tok = clean_tok.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
                            if col_tok and col_tok not in self._tag_to_category_order_collapsed:
                                self._tag_to_category_order_collapsed[col_tok] = order
                        en_lower = en.lower()
                        self._tag_to_category_order[en_lower] = order
                        col_en = en_lower.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
                        if col_en and col_en not in self._tag_to_category_order_collapsed:
                            self._tag_to_category_order_collapsed[col_en] = order

                # Common aliases / variations for camera angles
                cat2_order = cat_order_map.get(2, 2)
                for alias in ["bird eye's view", "birdeyesview", "bird's eye view", "birds eye view"]:
                    self._tag_to_category_order[alias] = cat2_order
                    self._tag_to_category_order_collapsed[alias.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")] = cat2_order
                for alias in ["worm eye's view", "wormeyesview", "worm's eye view", "worms eye view"]:
                    self._tag_to_category_order[alias] = cat2_order
                    self._tag_to_category_order_collapsed[alias.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")] = cat2_order

        except Exception as e:
            print(f"[DBManager] Error loading database: {e}")

        # Always load and merge user dictionary after loading tags.db
        self._load_and_merge_user_tags()

    def _load_and_merge_user_tags(self):
        """Loads user_tags.json and merges into in-memory category/tag caches."""
        self._user_tags_cache.clear()

        # 1. Ensure user category exists in categories cache
        user_cat = {
            "id": USER_CATEGORY_ID,
            "category_order": USER_CATEGORY_ORDER,
            "category_name": USER_CATEGORY_NAME
        }
        if not any(c["id"] == USER_CATEGORY_ID for c in self._categories_cache):
            self._categories_cache.append(user_cat)
            self._categories_cache.sort(key=lambda c: c.get("category_order", 999))

        self._tags_by_category_cache[USER_CATEGORY_ID] = []

        # 2. Read user_tags.json if exists
        if os.path.exists(self.user_tags_path):
            try:
                with open(self.user_tags_path, "r", encoding="utf-8-sig") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        self._user_tags_cache = data
            except Exception as e:
                print(f"[DBManager] Error reading user_tags.json: {e}")

        # 3. Merge each user tag into caches
        for t in self._user_tags_cache:
            tag_dict = dict(t)
            tag_dict["category_id"] = USER_CATEGORY_ID
            tag_dict["category_name"] = USER_CATEGORY_NAME
            if "note" not in tag_dict or tag_dict["note"] is None:
                tag_dict["note"] = ""

            self._tags_by_category_cache[USER_CATEGORY_ID].append(tag_dict)

            ja = tag_dict.get("label_ja", "").strip()
            en = tag_dict.get("prompt_en", "").strip()

            if ja and en:
                self._jp_to_en_map[ja] = en
                self._en_to_jp_map[en.lower()] = ja

            if en:
                order = USER_CATEGORY_ORDER
                for token in en.split(","):
                    clean_tok = token.strip().lower()
                    if clean_tok and clean_tok not in self._tag_to_category_order:
                        self._tag_to_category_order[clean_tok] = order
                    col_tok = clean_tok.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
                    if col_tok and col_tok not in self._tag_to_category_order_collapsed:
                        self._tag_to_category_order_collapsed[col_tok] = order
                en_lower = en.lower()
                self._tag_to_category_order[en_lower] = order
                col_en = en_lower.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
                if col_en and col_en not in self._tag_to_category_order_collapsed:
                    self._tag_to_category_order_collapsed[col_en] = order

    def get_user_tags(self) -> List[Dict]:
        """Returns all custom user tags."""
        return list(self._user_tags_cache)

    def save_user_tags(self) -> bool:
        """Saves current user_tags_cache to user_tags.json atomically."""
        try:
            tmp_path = self.user_tags_path + ".tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(self._user_tags_cache, f, indent=2, ensure_ascii=False)
            if os.path.exists(self.user_tags_path):
                os.replace(tmp_path, self.user_tags_path)
            else:
                os.rename(tmp_path, self.user_tags_path)
            return True
        except Exception as e:
            print(f"[DBManager] Error saving user_tags.json: {e}")
            return False

    def add_user_tag(self, label_ja: str, prompt_en: str, note: str = "") -> Tuple[bool, Optional[Dict], str]:
        """
        Adds a new tag to the user dictionary and caches.
        Returns (success: bool, new_tag_dict: Optional[Dict], message: str).
        """
        ja = label_ja.strip()
        en = prompt_en.strip()
        nt = note.strip()

        if not en:
            return False, None, "英語プロンプトが入力されていません。(Prompt is empty)"

        # Check duplicate
        for existing in self._user_tags_cache:
            if existing.get("prompt_en", "").strip().lower() == en.lower():
                return False, None, f"この英語プロンプトは既にユーザー辞書に登録されています: {en}"

        import time
        tag_id = f"u_{int(time.time() * 1000)}"
        new_tag = {
            "id": tag_id,
            "category_id": USER_CATEGORY_ID,
            "category_name": USER_CATEGORY_NAME,
            "label_ja": ja,
            "prompt_en": en,
            "note": nt or "User Dictionary"
        }

        self._user_tags_cache.append(new_tag)
        self._tags_by_category_cache[USER_CATEGORY_ID].append(new_tag)

        if ja:
            self._jp_to_en_map[ja] = en
            self._en_to_jp_map[en.lower()] = ja

        order = USER_CATEGORY_ORDER
        for token in en.split(","):
            clean_tok = token.strip().lower()
            if clean_tok and clean_tok not in self._tag_to_category_order:
                self._tag_to_category_order[clean_tok] = order
            col_tok = clean_tok.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
            if col_tok and col_tok not in self._tag_to_category_order_collapsed:
                self._tag_to_category_order_collapsed[col_tok] = order
        en_lower = en.lower()
        self._tag_to_category_order[en_lower] = order
        col_en = en_lower.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
        if col_en and col_en not in self._tag_to_category_order_collapsed:
            self._tag_to_category_order_collapsed[col_en] = order

        saved = self.save_user_tags()
        if not saved:
            return False, new_tag, "ユーザー辞書ファイルの保存に失敗しました。"

        return True, new_tag, f"「{ja or en}」をユーザー辞書に登録しました。"

    def delete_user_tag(self, tag_id: Any) -> bool:
        """Deletes a user tag by ID and updates caches and file."""
        target = None
        for t in self._user_tags_cache:
            if str(t.get("id")) == str(tag_id):
                target = t
                break

        if not target:
            return False

        self._user_tags_cache.remove(target)
        if USER_CATEGORY_ID in self._tags_by_category_cache:
            self._tags_by_category_cache[USER_CATEGORY_ID] = [
                t for t in self._tags_by_category_cache[USER_CATEGORY_ID]
                if str(t.get("id")) != str(tag_id)
            ]

        ja = target.get("label_ja", "").strip()
        en = target.get("prompt_en", "").strip().lower()
        if ja in self._jp_to_en_map and self._jp_to_en_map[ja] == target.get("prompt_en", "").strip():
            del self._jp_to_en_map[ja]
        if en in self._en_to_jp_map and self._en_to_jp_map[en] == ja:
            del self._en_to_jp_map[en]

        self.save_user_tags()
        return True

    def get_categories(self) -> List[Dict]:
        """Returns all categories sorted by category_order."""
        return self._categories_cache

    def get_tags_by_category(self, category_id: int) -> List[Dict]:
        """Returns all tags for a specific category ID."""
        return self._tags_by_category_cache.get(category_id, [])

    def search_tags(self, keyword: str) -> List[Dict]:
        """Searches tags by Japanese label, English prompt, or explanation Note."""
        kw = keyword.strip().lower()
        if not kw:
            return []

        kw_tokens = kw.replace(",", " ").split()
        results = []
        for cid, tag_list in self._tags_by_category_cache.items():
            for t in tag_list:
                ja = t.get("label_ja", "").lower()
                en = t.get("prompt_en", "").lower()
                note = str(t.get("note") or "").lower()
                # Match if all search query tokens appear in either ja, en, or note
                if all(tok in ja or tok in en or tok in note or 
                       tok.replace(" ", "") in ja.replace(" ", "") or 
                       tok.replace(" ", "") in en.replace(" ", "") or 
                       tok.replace(" ", "") in note.replace(" ", "") for tok in kw_tokens):
                    results.append(t)
        return results

    def get_tag_order(self, tag_en: str) -> int:
        """Returns the category order for an English tag (for prompt sorting)."""
        clean = tag_en.strip().lower()
        if clean in self._tag_to_category_order:
            return self._tag_to_category_order[clean]
        collapsed = clean.replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
        return self._tag_to_category_order_collapsed.get(collapsed, 999)

    def translate_jp_to_en(self, jp_label: str) -> Optional[str]:
        return self._jp_to_en_map.get(jp_label.strip())

    def translate_en_to_jp(self, en_prompt: str) -> Optional[str]:
        return self._en_to_jp_map.get(en_prompt.strip().lower())

    def is_known_tag(self, text: str) -> bool:
        """Checks if the given English text matches any known tag in the database (case- and whitespace-insensitive)."""
        if not text:
            return False
        import re
        norm = re.sub(r"\s+", " ", text.strip().lower()).replace(", ", ",").replace(",", ", ")
        if norm in self._en_to_jp_map:
            return True
        # Compare without spaces/underscores/hyphens/apostrophes
        raw_key = text.strip().lower().replace("'", "").replace("-", "").replace(" ", "").replace("_", "")
        if raw_key in self._tag_to_category_order_collapsed:
            return True
        for en_key in self._en_to_jp_map:
            if en_key.replace("'", "").replace("-", "").replace(" ", "").replace("_", "") == raw_key:
                return True
        return False

    def get_comma_tags(self) -> List[str]:
        """Returns all registered English tags that contain commas, sorted by length descending."""
        if not hasattr(self, "_comma_tags_cache") or not self._comma_tags_cache:
            self._comma_tags_cache = sorted(
                [en for en in self._en_to_jp_map.keys() if "," in en],
                key=len,
                reverse=True
            )
        return self._comma_tags_cache

    def get_sample_prompts(self) -> List[Dict[str, Any]]:
        """
        Retrieves sample prompts from SQLite 'sample_prompts' table (id, title, prompt, comment).
        If table does not exist or is empty, falls back to KENZEN_Sample_Prompts.json.
        """
        samples = []
        # 1. Try SQLite table
        if os.path.exists(self.db_path):
            try:
                with self.get_connection() as conn:
                    cur = conn.cursor()
                    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='sample_prompts';")
                    if cur.fetchone():
                        cur.execute("PRAGMA table_info(sample_prompts);")
                        cols = [c["name"] for c in cur.fetchall()]
                        has_comment = "comment" in cols
                        has_category = "category" in cols

                        col_query = ["id", "title", "prompt"]
                        if has_comment:
                            col_query.append("comment")
                        if has_category:
                            col_query.append("category")

                        cur.execute(f"SELECT {', '.join(col_query)} FROM sample_prompts ORDER BY id ASC;")
                        for row in cur.fetchall():
                            r_dict = dict(row)
                            if "comment" not in r_dict or r_dict["comment"] is None:
                                r_dict["comment"] = ""
                            if "category" not in r_dict or r_dict["category"] is None:
                                r_dict["category"] = "General"
                            samples.append(r_dict)
            except Exception as e:
                print(f"[DBManager] Note: Could not query sample_prompts table ({e}). Using JSON fallback.")

        if samples:
            return samples

        # 2. Fallback to KENZEN_Sample_Prompts.json
        import json
        from .path_utils import get_resource_path
        json_path = get_resource_path("KENZEN_Sample_Prompts.json")
        if os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8-sig") as f:
                    raw = json.load(f)
                    favs = raw.get("Favorites", [])
                    for idx, fv in enumerate(favs):
                        samples.append({
                            "id": fv.get("id", idx + 1),
                            "title": fv.get("description", f"Sample #{idx+1}"),
                            "prompt": fv.get("prompt", ""),
                            "comment": fv.get("comment", ""),
                            "category": "Sample"
                        })
            except Exception as e:
                print(f"[DBManager] Error reading sample prompts json: {e}")

        return samples
