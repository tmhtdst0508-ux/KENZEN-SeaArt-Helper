"""
Single Instance Manager for KENZEN SeaArt Helper v5.0.0
Uses Qt Local IPC (QLocalServer / QLocalSocket) to strictly prevent
multiple instances, showing a bilingual dialog and bringing the existing
window and Dictionary Matrix to the front.
"""

import sys
from PySide6.QtCore import QObject, Signal
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtWidgets import QMessageBox, QApplication


SERVER_NAME = "KENZEN_SeaArt_Helper_v5_SingleInstanceServer"


class SingleInstanceManager(QObject):
    activate_requested = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.server: QLocalServer = None

    def check_already_running(self) -> bool:
        """
        Attempts to connect to an existing running instance.
        If connection succeeds, notifies the existing instance, shows a warning dialog,
        and returns True (meaning duplicate instance).
        """
        socket = QLocalSocket()
        socket.connectToServer(SERVER_NAME)
        
        # If successfully connected within 500ms, an instance is already active
        if socket.waitForConnected(500):
            socket.write(b"ACTIVATE")
            socket.waitForBytesWritten(500)
            socket.disconnectFromServer()

            app = QApplication.instance() or QApplication(sys.argv)
            msg = (
                "【⚠️ 既に起動しています / Already Running】\n\n"
                "KENZEN SeaArt Helper は既に起動しています。\n"
                "起動中のウィンドウを最前面に表示しました。\n\n"
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                "KENZEN SeaArt Helper is already running.\n"
                "The active window has been brought to the front.\n"
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
            )
            box = QMessageBox()
            box.setIcon(QMessageBox.Information)
            box.setWindowTitle("多重起動の防止 / Already Running")
            box.setText(msg)
            box.setStandardButtons(QMessageBox.Ok)
            box.exec()
            return True

        # Clean up stale socket file if previously crashed
        QLocalServer.removeServer(SERVER_NAME)

        # Start server for this primary instance
        self.server = QLocalServer(self)
        self.server.newConnection.connect(self._on_new_connection)
        self.server.listen(SERVER_NAME)
        return False

    def _on_new_connection(self):
        """Called when a secondary instance attempts to launch."""
        client_socket = self.server.nextPendingConnection()
        if client_socket:
            self.activate_requested.emit()
            client_socket.disconnectFromServer()

    def cleanup(self):
        if self.server:
            self.server.close()
            QLocalServer.removeServer(SERVER_NAME)
