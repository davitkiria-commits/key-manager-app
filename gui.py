from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QIntValidator
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from key_manager import KeyManager, KeyManagerError


class AddApartmentDialog(QDialog):
    def __init__(self, manager: KeyManager, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Добавить квартиру")

        self.building_edit = QLineEdit()
        self.floor_edit = QLineEdit()
        self.floor_edit.setValidator(QIntValidator(-100, 300, self))
        self.apartment_edit = QLineEdit()
        self.total_keys_spin = QSpinBox()
        self.total_keys_spin.setRange(1, 1000)

        form = QFormLayout()
        form.addRow("Корпус:", self.building_edit)
        form.addRow("Этаж:", self.floor_edit)
        form.addRow("Квартира:", self.apartment_edit)
        form.addRow("Всего ключей:", self.total_keys_spin)

        save_btn = QPushButton("Добавить")
        save_btn.clicked.connect(self._save)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(save_btn)

    def _save(self) -> None:
        try:
            building = self.building_edit.text().strip()
            floor_text = self.floor_edit.text().strip()
            apt = self.apartment_edit.text().strip()
            total = self.total_keys_spin.value()

            if not building or not floor_text or not apt:
                raise KeyManagerError("Заполните все поля.")

            self.manager.add_apartment(building=building, floor=int(floor_text), apartment_number=apt, total_keys=total)
            self.accept()
        except (ValueError, KeyManagerError) as e:
            QMessageBox.warning(self, "Ошибка", str(e))


class IssueDialog(QDialog):
    def __init__(self, manager: KeyManager, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Выдача ключей")

        self.apartment_combo = QComboBox()
        self.recipient_combo = QComboBox()
        self.recipient_combo.setEditable(True)
        self.recipient_combo.setInsertPolicy(QComboBox.NoInsert)

        self.count_spin = QSpinBox()
        self.count_spin.setRange(1, 1000)

        self.limit_label = QLabel("Доступно: -")

        form = QFormLayout()
        form.addRow("Квартира:", self.apartment_combo)
        form.addRow("Получатель:", self.recipient_combo)
        form.addRow("Количество:", self.count_spin)
        form.addRow("Лимит:", self.limit_label)

        issue_btn = QPushButton("Выдать")
        issue_btn.clicked.connect(self._issue)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(issue_btn)

        self.apartment_combo.currentIndexChanged.connect(self._update_limit)
        self._reload_data()

    def _reload_data(self) -> None:
        self.apartment_combo.clear()
        for a in self.manager.get_apartments():
            self.apartment_combo.addItem(
                f"Корпус {a.building}, этаж {a.floor}, кв. {a.apartment_number}",
                a.apartment_id,
            )

        self.recipient_combo.clear()
        self.recipient_combo.addItems(self.manager.get_recipients())

        completer = self.recipient_combo.completer()
        if completer:
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)

        self._update_limit()

    def _update_limit(self) -> None:
        apartment_id = self.apartment_combo.currentData()
        apartment = self.manager.get_apartment(apartment_id) if apartment_id is not None else None
        if apartment:
            self.limit_label.setText(f"Доступно: {apartment.available_keys}")
            self.count_spin.setMaximum(max(1, apartment.available_keys))
        else:
            self.limit_label.setText("Доступно: -")

    def _issue(self) -> None:
        try:
            apartment_id = self.apartment_combo.currentData()
            if apartment_id is None:
                raise KeyManagerError("Нет доступных квартир для выдачи.")
            recipient = self.recipient_combo.currentText().strip()
            count = self.count_spin.value()
            self.manager.issue_keys(apartment_id=apartment_id, recipient=recipient, count=count)
            self.accept()
        except KeyManagerError as e:
            QMessageBox.warning(self, "Ошибка выдачи", str(e))


class ReturnDialog(QDialog):
    def __init__(self, manager: KeyManager, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Возврат ключей")

        self.issue_combo = QComboBox()
        self.count_spin = QSpinBox()
        self.count_spin.setRange(1, 1000)
        self.remaining_label = QLabel("Осталось вернуть: -")

        form = QFormLayout()
        form.addRow("Активная выдача:", self.issue_combo)
        form.addRow("Количество возврата:", self.count_spin)
        form.addRow("Статус:", self.remaining_label)

        return_btn = QPushButton("Вернуть")
        return_btn.clicked.connect(self._return_keys)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(return_btn)

        self.issue_combo.currentIndexChanged.connect(self._update_limit)
        self._reload_data()

    def _reload_data(self) -> None:
        self.issue_combo.clear()
        for issue in self.manager.get_active_issues():
            apartment = self.manager.get_apartment(issue.apartment_id)
            if apartment:
                text = (
                    f"#{issue.issue_id} | {issue.recipient} | "
                    f"кв. {apartment.apartment_number} | осталось: {issue.active_count}"
                )
                self.issue_combo.addItem(text, issue.issue_id)

        self._update_limit()

    def _update_limit(self) -> None:
        issue_id = self.issue_combo.currentData()
        issue = next((x for x in self.manager.get_active_issues() if x.issue_id == issue_id), None)
        if issue:
            self.remaining_label.setText(f"Осталось вернуть: {issue.active_count}")
            self.count_spin.setMaximum(max(1, issue.active_count))
        else:
            self.remaining_label.setText("Осталось вернуть: -")

    def _return_keys(self) -> None:
        try:
            issue_id = self.issue_combo.currentData()
            if issue_id is None:
                raise KeyManagerError("Нет активных выдач для возврата.")

            self.manager.return_keys(issue_id=issue_id, count=self.count_spin.value())
            self.accept()
        except KeyManagerError as e:
            QMessageBox.warning(self, "Ошибка возврата", str(e))


class LostDialog(QDialog):
    def __init__(self, manager: KeyManager, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Отметить утерю")

        self.apartment_combo = QComboBox()
        self.count_spin = QSpinBox()
        self.count_spin.setRange(1, 1000)
        self.limit_label = QLabel("Доступно: -")

        form = QFormLayout()
        form.addRow("Квартира:", self.apartment_combo)
        form.addRow("Количество:", self.count_spin)
        form.addRow("Лимит:", self.limit_label)

        lost_btn = QPushButton("Списать")
        lost_btn.clicked.connect(self._mark_lost)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(lost_btn)

        self.apartment_combo.currentIndexChanged.connect(self._update_limit)
        self._reload_data()

    def _reload_data(self) -> None:
        self.apartment_combo.clear()
        for a in self.manager.get_apartments():
            self.apartment_combo.addItem(
                f"Корпус {a.building}, этаж {a.floor}, кв. {a.apartment_number}",
                a.apartment_id,
            )
        self._update_limit()

    def _update_limit(self) -> None:
        apartment_id = self.apartment_combo.currentData()
        apartment = self.manager.get_apartment(apartment_id) if apartment_id is not None else None
        if apartment:
            self.limit_label.setText(f"Доступно: {apartment.available_keys}")
            self.count_spin.setMaximum(max(1, apartment.available_keys))
        else:
            self.limit_label.setText("Доступно: -")

    def _mark_lost(self) -> None:
        try:
            apartment_id = self.apartment_combo.currentData()
            if apartment_id is None:
                raise KeyManagerError("Нет доступных квартир.")

            self.manager.mark_lost(apartment_id=apartment_id, count=self.count_spin.value())
            self.accept()
        except KeyManagerError as e:
            QMessageBox.warning(self, "Ошибка", str(e))


class HistoryDialog(QDialog):
    def __init__(self, manager: KeyManager, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("История операций")
        self.resize(800, 500)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Время", "Тип", "Квартира", "Детали"])
        self.table.setSortingEnabled(True)

        layout = QVBoxLayout(self)
        layout.addWidget(self.table)

        self._fill()

    def _fill(self) -> None:
        history = self.manager.get_history()
        self.table.setRowCount(len(history))

        for row, item in enumerate(history):
            apartment = self.manager.get_apartment(item.apartment_id)
            apt_text = "-"
            if apartment:
                apt_text = f"{apartment.building}/{apartment.floor}/{apartment.apartment_number}"

            self.table.setItem(row, 0, QTableWidgetItem(item.timestamp.strftime("%Y-%m-%d %H:%M:%S")))
            self.table.setItem(row, 1, QTableWidgetItem(item.operation_type))
            self.table.setItem(row, 2, QTableWidgetItem(apt_text))
            self.table.setItem(row, 3, QTableWidgetItem(item.details))

        self.table.sortItems(0, Qt.DescendingOrder)


class MainWindow(QMainWindow):
    def __init__(self, manager: KeyManager) -> None:
        super().__init__()
        self.manager = manager
        self.setWindowTitle("Учет ключей")
        self.resize(1000, 600)

        root = QWidget()
        self.setCentralWidget(root)

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(
            ["ID", "Корпус", "Этаж", "Квартира", "Всего ключей", "Выдано", "Утеряно", "Доступно"]
        )
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        btn_add = QPushButton("Добавить квартиру")
        btn_issue = QPushButton("Выдать ключи")
        btn_return = QPushButton("Вернуть ключи")
        btn_lost = QPushButton("Отметить утерю")
        btn_history = QPushButton("История")

        btn_add.clicked.connect(self._on_add)
        btn_issue.clicked.connect(self._on_issue)
        btn_return.clicked.connect(self._on_return)
        btn_lost.clicked.connect(self._on_lost)
        btn_history.clicked.connect(self._on_history)

        btns = QHBoxLayout()
        for btn in (btn_add, btn_issue, btn_return, btn_lost, btn_history):
            btns.addWidget(btn)

        layout = QVBoxLayout(root)
        layout.addWidget(self.table)
        layout.addLayout(btns)

        self.refresh_table()

    def refresh_table(self) -> None:
        apartments = self.manager.get_apartments()
        self.table.setRowCount(len(apartments))

        for row, apt in enumerate(apartments):
            values = [
                str(apt.apartment_id),
                apt.building,
                str(apt.floor),
                apt.apartment_number,
                str(apt.total_keys),
                str(apt.issued_keys),
                str(apt.lost_keys),
                str(apt.available_keys),
            ]
            for col, value in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(value))

        self.table.resizeColumnsToContents()

    def _on_add(self) -> None:
        dialog = AddApartmentDialog(self.manager, self)
        if dialog.exec():
            self.refresh_table()

    def _on_issue(self) -> None:
        dialog = IssueDialog(self.manager, self)
        if dialog.exec():
            self.refresh_table()

    def _on_return(self) -> None:
        dialog = ReturnDialog(self.manager, self)
        if dialog.exec():
            self.refresh_table()

    def _on_lost(self) -> None:
        dialog = LostDialog(self.manager, self)
        if dialog.exec():
            self.refresh_table()

    def _on_history(self) -> None:
        dialog = HistoryDialog(self.manager, self)
        dialog.exec()
