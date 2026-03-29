import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from gui import MainWindow
from key_manager import KeyManager
from storage import StorageError, ensure_data_file_exists, load_data, save_data


def get_data_file_path() -> str:
    return str(get_app_dir() / "data.json")


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


CONFIG_FILE = get_app_dir() / "config.json"


def load_config_data_path() -> str:
    default_data_path = get_data_file_path()
    if not CONFIG_FILE.exists():
        return default_data_path

    config = load_data(str(CONFIG_FILE))
    data_path = config.get("data_path", default_data_path)
    if not isinstance(data_path, str) or not data_path.strip():
        raise StorageError("Некорректный config.json: поле 'data_path' должно быть непустой строкой.")
    return data_path.strip()


def save_config_data_path(data_path: str) -> None:
    save_data(str(CONFIG_FILE), {"data_path": data_path})


def main() -> int:
    app = QApplication(sys.argv)
    manager = KeyManager()

    try:
        data_file = load_config_data_path()
        ensure_data_file_exists(data_file)
    except StorageError as e:
        QMessageBox.warning(
            None,
            "Ошибка настроек",
            f"Не удалось применить config.json. Будет использован стандартный data.json.\n\n{e}",
        )
        data_file = get_data_file_path()
        try:
            ensure_data_file_exists(data_file)
        except StorageError as create_error:
            QMessageBox.critical(None, "Критическая ошибка", str(create_error))
            return 1

    def save_current_data() -> None:
        save_data(data_file, manager.to_dict())

    manager.set_on_change(save_current_data)

    try:
        manager.load_dict(load_data(data_file))
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

    def apply_data_path(new_data_path: str) -> None:
        nonlocal data_file
        new_data_path = new_data_path.strip()
        if not new_data_path:
            raise StorageError("Путь к data.json не может быть пустым.")

        ensure_data_file_exists(new_data_path)

        loaded_data = load_data(new_data_path)
        manager.load_dict(loaded_data)

        data_file = new_data_path
        save_config_data_path(new_data_path)
        save_current_data()

    window = MainWindow(manager, data_file, apply_data_path)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
