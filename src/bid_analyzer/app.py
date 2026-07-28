from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from bid_analyzer.gui.GUI_main_pyside6 import BidGUI


def main() -> int:
    """
    Start the desktop PySide6 application.
    """
    app = QApplication(sys.argv)

    window = BidGUI()
    window.show()

    return app.exec()