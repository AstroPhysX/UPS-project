"""UPS Bid Analyzer package entry point with a lightweight splash screen.

Place this file at:
    src/bid_analyzer/__main__.py

The splash is shown before the large GUI module is imported and remains visible
while the main window is constructed. It prefers app_icon.png for a sharp splash
and keeps app_icon.ico for the Windows application/window icon.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path


CONFIG_PATH = Path("bid_config.json")
DEFAULT_UI_SCALE_PERCENT = 100
UPS_BROWN = "#351C15"
UPS_GOLD = "#FFB500"
UPS_TEXT = "#FFF7E6"


def load_ui_scale_percent() -> int:
    try:
        with CONFIG_PATH.open("r", encoding="utf-8") as file:
            data = json.load(file)
        value = int(data.get("ui_scale_percent", DEFAULT_UI_SCALE_PERCENT))
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        value = DEFAULT_UI_SCALE_PERCENT

    return max(80, min(200, value))


def resource_candidates(filename: str) -> list[Path]:
    candidates = [
        Path(__file__).resolve().parent / "resources" / filename,
        Path(__file__).resolve().parent / filename,
    ]

    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root:
        root = Path(bundle_root)
        candidates.extend([
            root / "bid_analyzer" / "resources" / filename,
            root / "resources" / filename,
            root / filename,
        ])

    return candidates


def find_resource(filename: str) -> Path | None:
    for candidate in resource_candidates(filename):
        if candidate.exists():
            return candidate
    return None


def find_splash_image() -> Path | None:
    """Prefer a full-resolution PNG; use the ICO only as a last fallback."""
    for filename in (
        "app_icon.png",
        "app_logo.png",
        "logo.png",
        "app_icon.ico",
    ):
        path = find_resource(filename)
        if path is not None:
            return path
    return None


def create_splash_pixmap(scale_percent: int):
    from PySide6.QtCore import QRect, Qt
    from PySide6.QtGui import QColor, QFont, QPainter, QPen, QPixmap

    factor = scale_percent / 100.0
    width = max(420, round(520 * factor))
    height = max(235, round(290 * factor))
    pixmap = QPixmap(width, height)
    pixmap.fill(QColor(UPS_BROWN))

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)

    border_width = max(2, round(3 * factor))
    margin = max(8, round(10 * factor))
    painter.setPen(QPen(QColor(UPS_GOLD), border_width))
    painter.drawRoundedRect(
        margin,
        margin,
        width - margin * 2,
        height - margin * 2,
        max(10, round(16 * factor)),
        max(10, round(16 * factor)),
    )

    splash_image_path = find_splash_image()
    icon_bottom = round(height * 0.53)
    if splash_image_path is not None:
        icon = QPixmap(str(splash_image_path))
        if not icon.isNull():
            icon_size = max(70, round(105 * factor))
            icon = icon.scaled(
                icon_size,
                icon_size,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation,
            )
            icon_x = (width - icon.width()) // 2
            icon_y = max(margin * 2, round(35 * factor))
            painter.drawPixmap(icon_x, icon_y, icon)
            icon_bottom = icon_y + icon.height()

    title_font = QFont("Segoe UI")
    title_font.setBold(True)
    title_font.setPointSizeF(22.0 * factor)
    painter.setFont(title_font)
    painter.setPen(QColor(UPS_GOLD))
    painter.drawText(
        QRect(20, icon_bottom + round(10 * factor), width - 40, round(48 * factor)),
        Qt.AlignCenter,
        "UPS Bid Analyzer",
    )

    subtitle_font = QFont("Segoe UI")
    subtitle_font.setPointSizeF(10.5 * factor)
    painter.setFont(subtitle_font)
    painter.setPen(QColor(UPS_TEXT))
    painter.drawText(
        QRect(20, height - round(66 * factor), width - 40, round(28 * factor)),
        Qt.AlignCenter,
        "Loading the analyzer interface…",
    )

    painter.end()
    return pixmap


def main() -> int:
    from PySide6.QtCore import QCoreApplication, Qt
    from PySide6.QtGui import QIcon
    from PySide6.QtWidgets import QApplication, QSplashScreen

    scale_percent = load_ui_scale_percent()

    app = QApplication(sys.argv)
    app.setApplicationName("UPS Bid Analyzer")
    app.setApplicationDisplayName("UPS Bid Analyzer")
    app.setProperty("ups_ui_scale_percent", scale_percent)
    app.setProperty("ups_ui_scale_initialized", True)

    # Use the platform's normal font as the baseline, then apply the user's
    # additional analyzer-specific scale. This still preserves Qt's automatic
    # operating-system DPI handling.
    font = app.font()
    base_point_size = font.pointSizeF() if font.pointSizeF() > 0 else 10.0
    font.setPointSizeF(base_point_size * scale_percent / 100.0)
    app.setFont(font)

    icon_path = find_resource("app_icon.ico")
    if icon_path is not None:
        app.setWindowIcon(QIcon(str(icon_path)))

    splash = QSplashScreen(
        create_splash_pixmap(scale_percent),
        Qt.SplashScreen | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint,
    )
    splash.show()
    splash.showMessage(
        "  Preparing controls and saved preferences…",
        Qt.AlignLeft | Qt.AlignBottom,
        Qt.white,
    )
    QCoreApplication.processEvents()

    # Importing after the splash is visible covers the slowest part of Python
    # startup and avoids showing a partially constructed main window.
    from bid_analyzer.gui.GUI_main_pyside6 import BidGUI

    splash.showMessage(
        "  Building the main window…",
        Qt.AlignLeft | Qt.AlignBottom,
        Qt.white,
    )
    QCoreApplication.processEvents()

    window = BidGUI()
    window.show()
    splash.finish(window)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
