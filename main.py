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


def ask_switch_mode(new_path: str) -> str | None:
    dialog = QMessageBox()
    dialog.setIcon(QMessageBox.Question)
    dialog.setWindowTitle("Смена data.json")
    dialog.setText("Вы выбрали новый путь к data.json.")
    dialog.setInformativeText(
        f"Новый файл:\n{new_path}\n\n"
        "Выберите, как продолжить:\n"
        "• Использовать новый файл (пустой, если файла ещё нет)\n"
        "• Перенести текущие данные в новый файл"
    )

    use_new_btn = dialog.addButton("Использовать новый файл", QMessageBox.AcceptRole)
    migrate_btn = dialog.addButton("Перенести текущие данные в новый файл", QMessageBox.ActionRole)
    cancel_btn = dialog.addButton("Отмена", QMessageBox.RejectRole)

    dialog.exec()
    clicked = dialog.clickedButton()
    if clicked == use_new_btn:
        return "use_new"
    if clicked == migrate_btn:
        return "migrate"
    if clicked == cancel_btn:
        return None
    return None


def ask_existing_file_mode(new_path: str) -> str | None:
    dialog = QMessageBox()
    dialog.setIcon(QMessageBox.Warning)
    dialog.setWindowTitle("Файл уже существует")
    dialog.setText("По указанному пути уже есть data.json.")
    dialog.setInformativeText(
        f"Файл:\n{new_path}\n\n"
        "Выберите действие:\n"
        "• Загрузить существующий файл\n"
        "• Перезаписать его текущими данными"
    )

    load_btn = dialog.addButton("Загрузить существующий файл", QMessageBox.AcceptRole)
    overwrite_btn = dialog.addButton("Перезаписать его текущими данными", QMessageBox.DestructiveRole)
    cancel_btn = dialog.addButton("Отмена", QMessageBox.RejectRole)

    dialog.exec()
    clicked = dialog.clickedButton()
    if clicked == load_btn:
        return "load_existing"
    if clicked == overwrite_btn:
        return "overwrite_existing"
    if clicked == cancel_btn:
        return None
    return None


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

    def load_into_manager(file_path: str) -> None:
        loaded_data = load_data(file_path)
        try:
            manager.load_dict(loaded_data)
        except Exception as e:
            raise StorageError(f"Файл повреждён или имеет некорректный формат: {file_path}. ({e})") from e

    def switch_to_data_path(new_data_path: str) -> None:
        nonlocal data_file
        data_file = new_data_path
        save_config_data_path(new_data_path)
        save_current_data()
        window.on_data_source_changed(new_data_path)

    def apply_data_path(new_data_path: str) -> bool:
        new_data_path = new_data_path.strip()
        if not new_data_path:
            raise StorageError("Путь к data.json не может быть пустым.")
        if new_data_path == data_file:
            QMessageBox.information(None, "Настройки", "Указан текущий путь data.json. Изменений нет.")
            return True

        current_data = manager.to_dict()
        target_path = Path(new_data_path)

        if target_path.exists():
            existing_mode = ask_existing_file_mode(new_data_path)
            if existing_mode is None:
                return False

            if existing_mode == "load_existing":
                load_into_manager(new_data_path)
                switch_to_data_path(new_data_path)
                QMessageBox.information(None, "Готово", "Загружены данные из существующего файла.")
                return True

            save_data(new_data_path, current_data)
            switch_to_data_path(new_data_path)
            QMessageBox.information(None, "Готово", "Существующий файл перезаписан текущими данными.")
            return True

        switch_mode = ask_switch_mode(new_data_path)
        if switch_mode is None:
            return False

        if switch_mode == "migrate":
            save_data(new_data_path, current_data)
            switch_to_data_path(new_data_path)
            QMessageBox.information(None, "Готово", "Текущие данные перенесены в новый файл.")
            return True

        ensure_data_file_exists(new_data_path)
        load_into_manager(new_data_path)
        switch_to_data_path(new_data_path)
        QMessageBox.information(None, "Готово", "Подключен новый файл данных.")
        return True

    window = MainWindow(manager, data_file, apply_data_path)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
