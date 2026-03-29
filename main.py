import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from gui import MainWindow
from key_manager import KeyManager
from storage import StorageError, ensure_data_file_exists, load_data, save_data
from translations import tr


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
        raise StorageError(tr("invalid_config_data_path"))
    return data_path.strip()


def save_config_data_path(data_path: str) -> None:
    save_data(str(CONFIG_FILE), {"data_path": data_path})


def ask_switch_mode(new_path: str) -> str | None:
    dialog = QMessageBox()
    dialog.setIcon(QMessageBox.Question)
    dialog.setWindowTitle(tr("data_source_change"))
    dialog.setText(tr("new_data_path_selected"))
    dialog.setInformativeText(
        tr("switch_mode_details", path=new_path)
    )

    use_new_btn = dialog.addButton(tr("use_new_file"), QMessageBox.AcceptRole)
    migrate_btn = dialog.addButton(tr("migrate_data"), QMessageBox.ActionRole)
    cancel_btn = dialog.addButton(tr("cancel"), QMessageBox.RejectRole)

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
    dialog.setWindowTitle(tr("file_exists"))
    dialog.setText(tr("data_file_exists"))
    dialog.setInformativeText(
        tr("existing_file_details", path=new_path)
    )

    load_btn = dialog.addButton(tr("load_existing_file"), QMessageBox.AcceptRole)
    overwrite_btn = dialog.addButton(tr("overwrite_current_data"), QMessageBox.DestructiveRole)
    cancel_btn = dialog.addButton(tr("cancel"), QMessageBox.RejectRole)

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
            tr("settings_error"),
            tr("config_apply_failed", error=e),
        )
        data_file = get_data_file_path()
        try:
            ensure_data_file_exists(data_file)
        except StorageError as create_error:
            QMessageBox.critical(None, tr("critical_error"), str(create_error))
            return 1

    def save_current_data() -> None:
        save_data(data_file, manager.to_dict())

    manager.set_on_change(save_current_data)

    try:
        manager.load_dict(load_data(data_file))
    except StorageError as e:
        QMessageBox.warning(
            None,
            tr("data_load_error"),
            tr("empty_data_started", error=e),
        )
    except Exception as e:
        QMessageBox.warning(
            None,
            tr("data_load_error"),
            tr("invalid_saved_format", error=e),
        )

    def load_into_manager(file_path: str) -> None:
        loaded_data = load_data(file_path)
        try:
            manager.load_dict(loaded_data)
        except Exception as e:
            raise StorageError(tr("file_corrupted", file_path=file_path, error=e)) from e

    def switch_to_data_path(new_data_path: str) -> None:
        nonlocal data_file
        data_file = new_data_path
        save_config_data_path(new_data_path)
        save_current_data()
        window.on_data_source_changed(new_data_path)

    def apply_data_path(new_data_path: str) -> bool:
        new_data_path = new_data_path.strip()
        if not new_data_path:
            raise StorageError(tr("data_path_empty"))
        if new_data_path == data_file:
            QMessageBox.information(None, tr("settings"), tr("same_data_path"))
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
                QMessageBox.information(None, tr("done"), tr("loaded_existing_data"))
                return True

            save_data(new_data_path, current_data)
            switch_to_data_path(new_data_path)
            QMessageBox.information(None, tr("done"), tr("overwritten_existing_file"))
            return True

        switch_mode = ask_switch_mode(new_data_path)
        if switch_mode is None:
            return False

        if switch_mode == "migrate":
            save_data(new_data_path, current_data)
            switch_to_data_path(new_data_path)
            QMessageBox.information(None, tr("done"), tr("migrated_to_new_file"))
            return True

        ensure_data_file_exists(new_data_path)
        load_into_manager(new_data_path)
        switch_to_data_path(new_data_path)
        QMessageBox.information(None, tr("done"), tr("new_data_connected"))
        return True

    window = MainWindow(manager, data_file, apply_data_path)
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
