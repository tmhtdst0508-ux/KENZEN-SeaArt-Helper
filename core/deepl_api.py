"""
DeepL API Client for KENZEN SeaArt Helper v5.2.0
Provides Free/Pro endpoint auto-switching, error diagnostics,
and PySide6 QThread worker for smooth, non-blocking UI integration.
"""

import json
import requests
from typing import Tuple, Optional
from PySide6.QtCore import QThread, Signal


class DeepLAPI:
    FREE_ENDPOINT = "https://api-free.deepl.com/v2/translate"
    PRO_ENDPOINT = "https://api.deepl.com/v2/translate"

    # Browser URLs that automatically redirect to the user's localized language
    KEY_URL = "https://www.deepl.com/your-account/keys"
    PROMO_URL = "https://www.deepl.com/pro-api"

    def __init__(self, api_key: str = ""):
        self.api_key = api_key.strip()

    def set_api_key(self, api_key: str):
        self.api_key = api_key.strip()

    def get_endpoint(self) -> str:
        """
        DeepL Free API keys typically end with ':fx'.
        If ':fx' is present, use the free endpoint; otherwise default to pro.
        """
        if self.api_key.endswith(":fx"):
            return self.FREE_ENDPOINT
        return self.PRO_ENDPOINT

    def translate(self, text: str, target_lang: str = "EN", source_lang: Optional[str] = None) -> Tuple[bool, str]:
        """
        Synchronous translation request to DeepL API.
        If source_lang is None, DeepL automatically detects the source language.
        Returns (success: bool, translated_text_or_error_message: str).
        """
        clean_text = text.strip()
        if not clean_text:
            return False, "翻訳するテキストが空です。(Empty text)"

        if not self.api_key:
            return False, "DeepL APIキーが設定されていません。(API key is missing)"

        endpoint = self.get_endpoint()
        headers = {
            "Authorization": f"DeepL-Auth-Key {self.api_key}",
            "User-Agent": "KENZEN-SeaArt-Helper/5.2.0"
        }
        data = {
            "text": [clean_text],
            "target_lang": target_lang
        }
        if source_lang:
            data["source_lang"] = source_lang

        try:
            response = requests.post(endpoint, headers=headers, json=data, timeout=12.0)

            if response.status_code == 200:
                res_json = response.json()
                translations = res_json.get("translations", [])
                if translations and "text" in translations[0]:
                    return True, translations[0]["text"].strip()
                return False, "DeepLからの応答に翻訳データが含まれていませんでした。(No translation returned)"

            elif response.status_code == 403:
                return False, "認証に失敗しました。APIキーが正しいか、有効期限が切れていないか確認してください。(403 Forbidden: Invalid API Key)"
            elif response.status_code == 456:
                return False, "月間の翻訳文字数上限に達しました。(456 Quota Exceeded: Character limit reached)"
            elif response.status_code == 429:
                return False, "リクエスト数が多すぎます。少し待ってから再試行してください。(429 Too Many Requests)"
            else:
                error_detail = response.text[:200]
                return False, f"DeepL APIエラー (HTTP {response.status_code}): {error_detail}"

        except requests.exceptions.Timeout:
            return False, "通信がタイムアウトしました。インターネット接続を確認してください。(Timeout)"
        except requests.exceptions.ConnectionError:
            return False, "DeepLサーバーに接続できませんでした。ネットワーク接続を確認してください。(Connection Error)"
        except Exception as e:
            return False, f"翻訳処理中にエラーが発生しました: {e}"


class DeepLTranslateWorker(QThread):
    """Worker thread to run translation in background without blocking Qt UI."""
    translation_finished = Signal(str)  # Emits translated text
    translation_error = Signal(str)     # Emits user-friendly error message

    def __init__(self, api_client: DeepLAPI, text: str, parent=None):
        super().__init__(parent)
        self.client = api_client
        self.text = text

    def run(self):
        success, result = self.client.translate(self.text)
        if success:
            self.translation_finished.emit(result)
        else:
            self.translation_error.emit(result)
