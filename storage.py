from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict


class StorageError(Exception):
    """Ошибка чтения/записи данных."""


def load_data(file_path: str) -> Dict[str, Any]:
    path = Path(file_path)
    if not path.exists():
        return {}

    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        raise StorageError(f"Файл данных повреждён: {path}. Проверьте JSON. ({e})") from e
    except OSError as e:
        raise StorageError(f"Не удалось прочитать файл данных: {path}. ({e})") from e

    if not isinstance(data, dict):
        raise StorageError(f"Некорректный формат данных в {path}: ожидается JSON-объект.")

    return data


def save_data(file_path: str, data: Dict[str, Any]) -> None:
    path = Path(file_path)
    try:
        with path.open("w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except OSError as e:
        raise StorageError(f"Не удалось сохранить данные в {path}. ({e})") from e


def ensure_data_file_exists(file_path: str) -> None:
    path = Path(file_path)
    if path.exists():
        return

    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)
    except OSError as e:
        raise StorageError(f"Не удалось создать файл данных: {path}. ({e})") from e
