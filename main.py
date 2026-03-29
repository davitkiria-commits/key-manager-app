import sys

from PySide6.QtWidgets import QApplication

from gui import MainWindow
from key_manager import KeyManager


def main() -> int:
    app = QApplication(sys.argv)

    manager = KeyManager()
    window = MainWindow(manager)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
