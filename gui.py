from __future__ import annotations

from typing import Callable

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QIntValidator
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
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

from key_manager import KeyManager, KeyManagerError, Person


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


class PersonEditDialog(QDialog):
    def __init__(self, manager: KeyManager, person: Person | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.person = person

        self.setWindowTitle("Редактировать получателя" if person else "Добавить получателя")

        self.name_edit = QLineEdit(person.full_name if person else "")
        self.role_edit = QLineEdit(person.role if person else "")

        form = QFormLayout()
        form.addRow("ФИО:", self.name_edit)
        form.addRow("Роль:", self.role_edit)

        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self._save)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(save_btn)

    def _save(self) -> None:
        try:
            if self.person:
                self.manager.edit_person(
                    person_id=self.person.person_id,
                    full_name=self.name_edit.text(),
                    role=self.role_edit.text(),
                )
            else:
                self.manager.add_person(
                    full_name=self.name_edit.text(),
                    role=self.role_edit.text(),
                    is_active=True,
                )
            self.accept()
        except KeyManagerError as e:
            QMessageBox.warning(self, "Ошибка", str(e))


class PersonsDialog(QDialog):
    def __init__(self, manager: KeyManager, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Получатели")
        self.resize(760, 500)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск по имени...")

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["ID", "ФИО", "Роль", "Активен"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        btn_add = QPushButton("Добавить")
        btn_edit = QPushButton("Редактировать")
        btn_toggle = QPushButton("Отключить/Активировать")

        btn_add.clicked.connect(self._on_add)
        btn_edit.clicked.connect(self._on_edit)
        btn_toggle.clicked.connect(self._on_toggle)
        self.search_edit.textChanged.connect(self._fill)

        controls = QHBoxLayout()
        controls.addWidget(btn_add)
        controls.addWidget(btn_edit)
        controls.addWidget(btn_toggle)

        layout = QVBoxLayout(self)
        layout.addWidget(self.search_edit)
        layout.addWidget(self.table)
        layout.addLayout(controls)

        self._fill()

    def _fill(self) -> None:
        query = self.search_edit.text().strip().lower()
        persons = self.manager.get_persons(include_inactive=True)
        if query:
            persons = [p for p in persons if query in p.full_name.lower()]

        self.table.setRowCount(len(persons))
        for row, person in enumerate(persons):
            self.table.setItem(row, 0, QTableWidgetItem(str(person.person_id)))
            self.table.setItem(row, 1, QTableWidgetItem(person.full_name))
            self.table.setItem(row, 2, QTableWidgetItem(person.role))
            self.table.setItem(row, 3, QTableWidgetItem("Да" if person.is_active else "Нет"))

        self.table.resizeColumnsToContents()

    def _selected_person_id(self) -> int | None:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 0)
        if not item:
            return None
        return int(item.text())

    def _on_add(self) -> None:
        dialog = PersonEditDialog(self.manager, parent=self)
        if dialog.exec():
            self._fill()

    def _on_edit(self) -> None:
        person_id = self._selected_person_id()
        if person_id is None:
            QMessageBox.information(self, "Внимание", "Выберите человека в таблице.")
            return

        person = self.manager.get_person(person_id)
        if not person:
            return
        dialog = PersonEditDialog(self.manager, person=person, parent=self)
        if dialog.exec():
            self._fill()

    def _on_toggle(self) -> None:
        person_id = self._selected_person_id()
        if person_id is None:
            QMessageBox.information(self, "Внимание", "Выберите человека в таблице.")
            return

        person = self.manager.get_person(person_id)
        if not person:
            return

        self.manager.set_person_active(person_id, not person.is_active)
        self._fill()


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
        self.person_hint_label = QLabel("")

        form = QFormLayout()
        form.addRow("Квартира:", self.apartment_combo)
        form.addRow("Получатель:", self.recipient_combo)
        form.addRow("", self.person_hint_label)
        form.addRow("Количество:", self.count_spin)
        form.addRow("Лимит:", self.limit_label)

        self.issue_btn = QPushButton("Выдать")
        self.issue_btn.clicked.connect(self._issue)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(self.issue_btn)

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
        active_persons = self.manager.get_persons(include_inactive=False)
        for person in active_persons:
            self.recipient_combo.addItem(f"{person.full_name} ({person.role})" if person.role else person.full_name, person.person_id)

        completer = self.recipient_combo.completer()
        if completer:
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)

        if not active_persons:
            self.person_hint_label.setText("Нет активных получателей. Добавьте людей в окне 'Получатели'.")
            self.recipient_combo.setEnabled(False)
            self.issue_btn.setEnabled(False)
        else:
            self.person_hint_label.setText("Показываются только активные получатели.")
            self.recipient_combo.setEnabled(True)
            self.issue_btn.setEnabled(True)

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

            person_id = self.recipient_combo.currentData()
            if person_id is None:
                raise KeyManagerError("Выберите получателя из списка.")

            count = self.count_spin.value()
            self.manager.issue_keys(apartment_id=apartment_id, recipient_id=person_id, count=count)
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
                    f"#{issue.issue_id} | {issue.recipient_name} | "
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
        self.resize(1100, 600)

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

        self.person_combo = QComboBox()
        self.person_combo.setEditable(True)
        self.person_combo.addItem("", "")
        for person in self.manager.get_persons(include_inactive=True):
            self.person_combo.addItem(person.full_name, person.full_name)
        completer = self.person_combo.completer()
        if completer:
            completer.setFilterMode(Qt.MatchContains)
            completer.setCaseSensitivity(Qt.CaseInsensitive)

        self.apartment_edit = QLineEdit()
        self.apartment_edit.setPlaceholderText("Номер квартиры")

        self.reset_btn = QPushButton("Сбросить фильтры")
        self.reset_btn.clicked.connect(self._reset_filters)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Дата с:"))
        filters.addWidget(self.date_from_edit)
        filters.addWidget(QLabel("Дата по:"))
        filters.addWidget(self.date_to_edit)
        filters.addWidget(QLabel("Человек:"))
        filters.addWidget(self.person_combo)
        filters.addWidget(QLabel("Квартира:"))
        filters.addWidget(self.apartment_edit)
        filters.addWidget(self.reset_btn)

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(
            ["Дата и время", "Действие", "Корпус", "Этаж", "Квартира", "Получатель", "Количество", "Статус"]
        )
        self.table.setSortingEnabled(True)

        layout = QVBoxLayout(self)
        layout.addLayout(filters)
        layout.addWidget(self.table)

        self.person_combo.currentIndexChanged.connect(self._fill)
        self.person_combo.lineEdit().textChanged.connect(self._fill)
        self.apartment_edit.textChanged.connect(self._fill)
        self.date_from_edit.dateChanged.connect(self._fill)
        self.date_to_edit.dateChanged.connect(self._fill)

        self._fill()

    def _reset_filters(self) -> None:
        self.person_combo.setCurrentIndex(0)
        self.person_combo.lineEdit().clear()
        self.apartment_edit.clear()
        self.date_from_edit.setDate(self.date_from_edit.minimumDate())
        self.date_to_edit.setDate(self.date_to_edit.minimumDate())
        self._fill()

    def _fill(self) -> None:
        person_text = self.person_combo.currentText().strip().lower()
        apt_text = self.apartment_edit.text().strip().lower()

        date_from = self.date_from_edit.date()
        date_to = self.date_to_edit.date()
        has_date_from = date_from != self.date_from_edit.minimumDate()
        has_date_to = date_to != self.date_to_edit.minimumDate()

        rows = []
        for item in self.manager.get_history():
            if person_text and person_text not in item.recipient.lower():
                continue
            if apt_text and apt_text not in str(item.apartment_number).lower():
                continue

            op_date = QDate(item.timestamp.year, item.timestamp.month, item.timestamp.day)
            if has_date_from and op_date < date_from:
                continue
            if has_date_to and op_date > date_to:
                continue

            rows.append(item)

        self.table.setRowCount(len(rows))
        for row, item in enumerate(rows):
            values = [
                item.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                item.operation_type,
                item.building,
                item.floor,
                item.apartment_number,
                item.recipient,
                str(item.quantity),
                item.status,
            ]
            for col, value in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(value))

        self.table.resizeColumnsToContents()
        self.table.sortItems(0, Qt.DescendingOrder)


class KeysOnHandDialog(QDialog):
    def __init__(self, manager: KeyManager, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.manager = manager
        self.setWindowTitle("Ключи на руках")
        self.resize(900, 500)

        self.table = QTableWidget(0, 6)
        self.table.setHorizontalHeaderLabels(["Корпус", "Этаж", "Квартира", "Кто держит", "Количество", "Дата выдачи"])

        layout = QVBoxLayout(self)
        layout.addWidget(self.table)

        self._fill()

    def _fill(self) -> None:
        issues = self.manager.get_active_issues()
        self.table.setRowCount(len(issues))

        for row, issue in enumerate(issues):
            apartment = self.manager.get_apartment(issue.apartment_id)
            if apartment:
                values = [
                    apartment.building,
                    str(apartment.floor),
                    apartment.apartment_number,
                    issue.recipient_name,
                    str(issue.active_count),
                    issue.issued_at.strftime("%Y-%m-%d %H:%M:%S"),
                ]
            else:
                values = ["-", "-", "-", issue.recipient_name, str(issue.active_count), issue.issued_at.strftime("%Y-%m-%d %H:%M:%S")]

            for col, value in enumerate(values):
                self.table.setItem(row, col, QTableWidgetItem(value))

        self.table.resizeColumnsToContents()


class SettingsDialog(QDialog):
    def __init__(
        self,
        current_data_path: str,
        on_save_path: Callable[[str], None],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._on_save_path = on_save_path
        self.setWindowTitle("Настройки")
        self.resize(650, 160)

        self.path_edit = QLineEdit(current_data_path)

        choose_btn = QPushButton("Выбрать файл")
        save_btn = QPushButton("Сохранить")
        choose_btn.clicked.connect(self._choose_file)
        save_btn.clicked.connect(self._save)

        form = QFormLayout()
        form.addRow("Путь к data.json:", self.path_edit)

        controls = QHBoxLayout()
        controls.addWidget(choose_btn)
        controls.addWidget(save_btn)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addLayout(controls)

    def _choose_file(self) -> None:
        selected_path, _ = QFileDialog.getSaveFileName(
            self,
            "Выберите data.json",
            self.path_edit.text().strip() or "data.json",
            "JSON Files (*.json);;All Files (*)",
        )
        if selected_path:
            self.path_edit.setText(selected_path)

    def _save(self) -> None:
        path = self.path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Ошибка", "Укажите путь к файлу data.json.")
            return
        try:
            self._on_save_path(path)
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Ошибка сохранения настроек", str(e))


class MainWindow(QMainWindow):
    def __init__(
        self,
        manager: KeyManager,
        data_file_path: str,
        on_change_data_path: Callable[[str], None],
    ) -> None:
        super().__init__()
        self.manager = manager
        self.data_file_path = data_file_path
        self.on_change_data_path = on_change_data_path
        self.setWindowTitle("Учет ключей")
        self.resize(1100, 650)

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
        btn_persons = QPushButton("Получатели")
        btn_keys_on_hand = QPushButton("Ключи на руках")
        btn_export_excel = QPushButton("Экспорт в Excel")
        btn_settings = QPushButton("Настройки")

        btn_add.clicked.connect(self._on_add)
        btn_issue.clicked.connect(self._on_issue)
        btn_return.clicked.connect(self._on_return)
        btn_lost.clicked.connect(self._on_lost)
        btn_history.clicked.connect(self._on_history)
        btn_persons.clicked.connect(self._on_persons)
        btn_keys_on_hand.clicked.connect(self._on_keys_on_hand)
        btn_export_excel.clicked.connect(self._on_export_excel)
        btn_settings.clicked.connect(self._on_settings)
        self.search_edit.textChanged.connect(self.refresh_table)

        btns = QHBoxLayout()
        for btn in (
            btn_add,
            btn_issue,
            btn_return,
            btn_lost,
            btn_history,
            btn_persons,
            btn_keys_on_hand,
            btn_export_excel,
            btn_settings,
        ):
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

    def _on_persons(self) -> None:
        dialog = PersonsDialog(self.manager, self)
        dialog.exec()

    def _on_keys_on_hand(self) -> None:
        dialog = KeysOnHandDialog(self.manager, self)
        dialog.exec()

    def _on_export_excel(self) -> None:
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить Excel-файл",
            "key_manager_export.xlsx",
            "Excel Files (*.xlsx)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            from openpyxl import Workbook
        except ImportError:
            QMessageBox.critical(
                self,
                "Ошибка экспорта",
                "Библиотека openpyxl не установлена. Установите зависимости из requirements.txt.",
            )
            return

        try:
            workbook = Workbook()

            apartments_sheet = workbook.active
            apartments_sheet.title = "Квартиры"
            apartments_sheet.append(["Корпус", "Этаж", "Квартира", "Всего ключей", "Выдано", "Утеряно", "Доступно"])
            for apartment in self.manager.get_apartments():
                apartments_sheet.append(
                    [
                        apartment.building,
                        apartment.floor,
                        apartment.apartment_number,
                        apartment.total_keys,
                        apartment.issued_keys,
                        apartment.lost_keys,
                        apartment.available_keys,
                    ]
                )

            history_sheet = workbook.create_sheet("История")
            history_sheet.append(["Дата/время", "Действие", "Корпус", "Этаж", "Квартира", "Получатель", "Количество", "Статус"])
            for item in self.manager.get_history():
                history_sheet.append(
                    [
                        item.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                        item.operation_type,
                        item.building,
                        item.floor,
                        item.apartment_number,
                        item.recipient,
                        item.quantity,
                        item.status,
                    ]
                )

            active_sheet = workbook.create_sheet("Ключи на руках")
            active_sheet.append(["Корпус", "Этаж", "Квартира", "Получатель", "Количество", "Дата выдачи"])
            for issue in self.manager.get_active_issues():
                apartment = self.manager.get_apartment(issue.apartment_id)
                if apartment:
                    row = [
                        apartment.building,
                        apartment.floor,
                        apartment.apartment_number,
                        issue.recipient_name,
                        issue.active_count,
                        issue.issued_at.strftime("%Y-%m-%d %H:%M:%S"),
                    ]
                else:
                    row = ["-", "-", "-", issue.recipient_name, issue.active_count, issue.issued_at.strftime("%Y-%m-%d %H:%M:%S")]
                active_sheet.append(row)

            workbook.save(file_path)
            QMessageBox.information(self, "Экспорт завершен", f"Данные успешно сохранены в файл:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка экспорта", f"Не удалось сохранить файл Excel.\n\n{e}")

    def _on_settings(self) -> None:
        dialog = SettingsDialog(self.data_file_path, self._apply_data_path, self)
        if dialog.exec():
            self.refresh_table()

    def _apply_data_path(self, path: str) -> None:
        self.on_change_data_path(path)
        self.data_file_path = path
