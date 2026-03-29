import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from gui import MainWindow
from key_manager import KeyManager
from storage import StorageError, load_data, save_data


def get_data_file_path() -> str:
    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent
    return str(base_dir / "data.json")


DATA_FILE = get_data_file_path()


def main() -> int:
    app = QApplication(sys.argv)

    manager = KeyManager(on_change=lambda: save_data(DATA_FILE, manager.to_dict()))

    try:
        manager.load_dict(load_data(DATA_FILE))
    except StorageError as e:
        QMessageBox.warning(
            None,
            "Ошибка загрузки данных",
            f"Программа запущена с пустыми данными.\n\n{e}",
        )
    except Exception as e:
        QMessageBox.warning(
            None,
            "Ошибка загрузки данных",
            f"Некорректный формат сохраненных данных. Программа запущена с пустыми данными.\n\n{e}",
        )

    window = MainWindow(manager)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
