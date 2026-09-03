from PySide6.QtWidgets import QTextEdit, QLineEdit
from PySide6.QtCore import Qt

class PlainTextOnlyTextEdit(QTextEdit):
    """
    Custom QTextEdit that strictly enforces plain text input.
    Strips all rich text formatting, styles, fonts, colors, and HTML tags
    when pasted via Ctrl+V, Shift+Insert, right-click context menu,
    or drag-and-drop from browsers and other rich applications.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setAcceptRichText(False)

    def insertFromMimeData(self, source):
        if source and source.hasText():
            self.insertPlainText(source.text())
        elif source:
            super().insertFromMimeData(source)


class PlainTextLineEdit(QLineEdit):
    """
    Custom QLineEdit that strictly enforces plain text input,
    stripping rich formats and normalizing any multiline text to single-line.
    """
    def insertFromMimeData(self, source):
        if source and source.hasText():
            clean_text = source.text().replace("\r\n", " ").replace("\n", " ")
            self.insert(clean_text)
        elif source:
            super().insertFromMimeData(source)
