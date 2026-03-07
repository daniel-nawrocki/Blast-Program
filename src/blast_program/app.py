import sys

from PySide6.QtWidgets import QApplication

from blast_program.ui.main_window import MainWindow


def run() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    try:
        return app.exec()
    except KeyboardInterrupt:
        return 0
