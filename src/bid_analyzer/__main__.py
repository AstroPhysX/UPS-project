"""
Package entry point for UPS Bid Analyzer.

Run from the project environment with:

    python -m bid_analyzer
"""

from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from bid_analyzer.gui.GUI_main_pyside6 import BidGUI


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("UPS Bid Analyzer")

    window = BidGUI()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
