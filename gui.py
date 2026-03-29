from __future__ import annotations

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QIntValidator
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
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
        self.resize(900, 550)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск по получателю...")

        self.date_from_edit = QDateEdit()
        self.date_from_edit.setCalendarPopup(True)
        self.date_from_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_from_edit.setSpecialValueText("-")
        self.date_from_edit.setMinimumDate(QDate(1900, 1, 1))
        self.date_from_edit.setDate(self.date_from_edit.minimumDate())

        self.date_to_edit = QDateEdit()
        self.date_to_edit.setCalendarPopup(True)
        self.date_to_edit.setDisplayFormat("yyyy-MM-dd")
        self.date_to_edit.setSpecialValueText("-")
        self.date_to_edit.setMinimumDate(QDate(1900, 1, 1))
        self.date_to_edit.setDate(self.date_to_edit.minimumDate())

        self.employee_edit = QLineEdit()
        self.employee_edit.setPlaceholderText("Сотрудник / получатель")

        self.reset_btn = QPushButton("Сбросить фильтры")
        self.reset_btn.clicked.connect(self._reset_filters)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Поиск получателя:"))
        filters.addWidget(self.search_edit)
        filters.addWidget(QLabel("Дата с:"))
        filters.addWidget(self.date_from_edit)
        filters.addWidget(QLabel("Дата по:"))
        filters.addWidget(self.date_to_edit)
        filters.addWidget(QLabel("Сотрудник:"))
        filters.addWidget(self.employee_edit)
        filters.addWidget(self.reset_btn)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Время", "Тип", "Квартира", "Детали"])
        self.table.setSortingEnabled(True)

        layout = QVBoxLayout(self)
        layout.addLayout(filters)
        layout.addWidget(self.table)

        self.search_edit.textChanged.connect(self._fill)
        self.employee_edit.textChanged.connect(self._fill)
        self.date_from_edit.dateChanged.connect(self._fill)
        self.date_to_edit.dateChanged.connect(self._fill)

        self._fill()

    def _reset_filters(self) -> None:
        self.search_edit.clear()
        self.employee_edit.clear()
        self.date_from_edit.setDate(self.date_from_edit.minimumDate())
        self.date_to_edit.setDate(self.date_to_edit.minimumDate())
        self._fill()

    def _fill(self) -> None:
        search_text = self.search_edit.text().strip().lower()
        employee_text = self.employee_edit.text().strip().lower()

        date_from = self.date_from_edit.date()
        date_to = self.date_to_edit.date()
        has_date_from = date_from != self.date_from_edit.minimumDate()
        has_date_to = date_to != self.date_to_edit.minimumDate()

        filtered = []
        for item in self.manager.get_history():
            recipient_text = item.recipient.lower() if item.recipient else ""
            details_text = item.details.lower()

            if search_text and search_text not in recipient_text and search_text not in details_text:
                continue
            if employee_text and employee_text not in recipient_text and employee_text not in details_text:
                continue

            op_date = QDate(item.timestamp.year, item.timestamp.month, item.timestamp.day)
            if has_date_from and op_date < date_from:
                continue
            if has_date_to and op_date > date_to:
                continue

            filtered.append(item)

        self.table.setRowCount(len(filtered))

        for row, item in enumerate(filtered):
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

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск по корпусу / этажу / квартире...")

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
        self.search_edit.textChanged.connect(self.refresh_table)

        btns = QHBoxLayout()
        for btn in (btn_add, btn_issue, btn_return, btn_lost, btn_history):
            btns.addWidget(btn)

        layout = QVBoxLayout(root)
        layout.addWidget(self.search_edit)
        layout.addWidget(self.table)
        layout.addLayout(btns)

        self.refresh_table()

    def refresh_table(self) -> None:
        apartments = self.manager.get_apartments()
        query = self.search_edit.text().strip().lower()

        if query:
            apartments = [
                apt
                for apt in apartments
                if query in apt.building.lower()
                or query in str(apt.floor).lower()
                or query in apt.apartment_number.lower()
            ]

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
