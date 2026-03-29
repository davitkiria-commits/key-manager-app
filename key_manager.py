from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Callable, Dict, List, Optional


@dataclass
class Apartment:
    apartment_id: int
    building: str
    floor: int
    apartment_number: str
    total_keys: int
    issued_keys: int = 0
    lost_keys: int = 0

    @property
    def available_keys(self) -> int:
        return self.total_keys - self.issued_keys - self.lost_keys


@dataclass
class Person:
    person_id: int
    full_name: str
    role: str
    is_active: bool = True


@dataclass
class ActiveIssue:
    issue_id: int
    apartment_id: int
    recipient_id: Optional[int]
    recipient_name: str
    issued_count: int
    returned_count: int = 0
    issued_at: datetime = field(default_factory=datetime.now)

    @property
    def active_count(self) -> int:
        return self.issued_count - self.returned_count


@dataclass
class Operation:
    timestamp: datetime
    operation_type: str
    apartment_id: Optional[int]
    building: str
    floor: str
    apartment_number: str
    recipient: str
    quantity: int
    status: str
    details: str


class KeyManagerError(Exception):
    """Base class for key manager errors."""


class KeyManager:
    def __init__(self, on_change: Callable[[], None] | None = None) -> None:
        self._apartments: Dict[int, Apartment] = {}
        self._active_issues: Dict[int, ActiveIssue] = {}
        self._history: List[Operation] = []
        self._persons: Dict[int, Person] = {}

        self._next_apartment_id = 1
        self._next_issue_id = 1
        self._next_person_id = 1
        self._on_change = on_change

    def set_on_change(self, on_change: Callable[[], None] | None) -> None:
        self._on_change = on_change

    def add_apartment(self, building: str, floor: int, apartment_number: str, total_keys: int) -> Apartment:
        if total_keys <= 0:
            raise KeyManagerError("Общее количество ключей должно быть больше 0.")

        apartment = Apartment(
            apartment_id=self._next_apartment_id,
            building=building.strip(),
            floor=floor,
            apartment_number=apartment_number.strip(),
            total_keys=total_keys,
        )
        self._apartments[apartment.apartment_id] = apartment
        self._next_apartment_id += 1

        self._add_history(
            operation_type="ADD_APARTMENT",
            apartment=apartment,
            recipient="",
            quantity=total_keys,
            status="Создана",
            details=f"Добавлена квартира, всего ключей: {total_keys}",
        )
        self._notify_change()
        return apartment

    def add_person(self, full_name: str, role: str, is_active: bool = True) -> Person:
        name = full_name.strip()
        if not name:
            raise KeyManagerError("Укажите ФИО.")

        person = Person(
            person_id=self._next_person_id,
            full_name=name,
            role=role.strip(),
            is_active=is_active,
        )
        self._persons[person.person_id] = person
        self._next_person_id += 1
        self._notify_change()
        return person

    def edit_person(self, person_id: int, full_name: str, role: str) -> None:
        person = self._persons.get(person_id)
        if not person:
            raise KeyManagerError("Получатель не найден.")
        name = full_name.strip()
        if not name:
            raise KeyManagerError("Укажите ФИО.")

        person.full_name = name
        person.role = role.strip()
        self._notify_change()

    def set_person_active(self, person_id: int, is_active: bool) -> None:
        person = self._persons.get(person_id)
        if not person:
            raise KeyManagerError("Получатель не найден.")
        person.is_active = is_active
        self._notify_change()

    def get_person(self, person_id: int) -> Optional[Person]:
        return self._persons.get(person_id)

    def get_persons(self, include_inactive: bool = True) -> List[Person]:
        persons = self._persons.values()
        if not include_inactive:
            persons = [p for p in persons if p.is_active]
        return sorted(persons, key=lambda p: p.full_name.lower())

    def issue_keys(self, apartment_id: int, recipient_id: int, count: int) -> ActiveIssue:
        apartment = self._get_apartment(apartment_id)

        if count <= 0:
            raise KeyManagerError("Количество выдаваемых ключей должно быть больше 0.")
        if count > apartment.available_keys:
            raise KeyManagerError(
                f"Недостаточно доступных ключей. Доступно: {apartment.available_keys}, запрошено: {count}."
            )

        person = self._persons.get(recipient_id)
        if not person:
            raise KeyManagerError("Выберите получателя из справочника.")
        if not person.is_active:
            raise KeyManagerError("Нельзя выдавать ключи неактивному получателю.")

        apartment.issued_keys += count

        issue = ActiveIssue(
            issue_id=self._next_issue_id,
            apartment_id=apartment_id,
            recipient_id=person.person_id,
            recipient_name=person.full_name,
            issued_count=count,
        )
        self._active_issues[issue.issue_id] = issue
        self._next_issue_id += 1

        self._add_history(
            operation_type="ISSUE",
            apartment=apartment,
            recipient=person.full_name,
            quantity=count,
            status="Выдано",
            details=f"Выдано {count} ключ(ей) получателю: {person.full_name}",
        )
        self._notify_change()
        return issue

    def return_keys(self, issue_id: int, count: int) -> None:
        issue = self._active_issues.get(issue_id)
        if not issue:
            raise KeyManagerError("Активная выдача не найдена.")
        if count <= 0:
            raise KeyManagerError("Количество возврата должно быть больше 0.")
        if count > issue.active_count:
            raise KeyManagerError(
                f"Нельзя вернуть больше, чем выдано. Осталось вернуть: {issue.active_count}, запрошено: {count}."
            )

        apartment = self._get_apartment(issue.apartment_id)
        issue.returned_count += count
        apartment.issued_keys -= count

        status = "Закрыто" if issue.active_count - count == 0 else "Частично возвращено"
        self._add_history(
            operation_type="RETURN",
            apartment=apartment,
            recipient=issue.recipient_name,
            quantity=count,
            status=status,
            details=f"Возвращено {count} ключ(ей) от получателя: {issue.recipient_name}",
        )

        if issue.active_count == 0:
            del self._active_issues[issue_id]

        self._notify_change()

    def mark_lost(self, apartment_id: int, count: int) -> None:
        apartment = self._get_apartment(apartment_id)

        if count <= 0:
            raise KeyManagerError("Количество утерянных ключей должно быть больше 0.")
        if count > apartment.available_keys:
            raise KeyManagerError(
                f"Недостаточно доступных ключей для списания. Доступно: {apartment.available_keys}, запрошено: {count}."
            )

        apartment.lost_keys += count
        self._add_history(
            operation_type="LOST",
            apartment=apartment,
            recipient="",
            quantity=count,
            status="Утеря",
            details=f"Отмечена утеря {count} ключ(ей)",
        )
        self._notify_change()

    def get_apartments(self) -> List[Apartment]:
        return sorted(self._apartments.values(), key=lambda a: (a.building, a.floor, a.apartment_number))

    def get_apartment(self, apartment_id: int) -> Optional[Apartment]:
        return self._apartments.get(apartment_id)

    def get_active_issues(self) -> List[ActiveIssue]:
        return sorted(self._active_issues.values(), key=lambda i: i.issued_at, reverse=True)

    def get_history(self) -> List[Operation]:
        return sorted(self._history, key=lambda h: h.timestamp, reverse=True)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "apartments": [
                {
                    "apartment_id": a.apartment_id,
                    "building": a.building,
                    "floor": a.floor,
                    "apartment_number": a.apartment_number,
                    "total_keys": a.total_keys,
                    "issued_keys": a.issued_keys,
                    "lost_keys": a.lost_keys,
                }
                for a in self._apartments.values()
            ],
            "persons": [
                {
                    "person_id": p.person_id,
                    "full_name": p.full_name,
                    "role": p.role,
                    "is_active": p.is_active,
                }
                for p in self._persons.values()
            ],
            "active_issues": [
                {
                    "issue_id": issue.issue_id,
                    "apartment_id": issue.apartment_id,
                    "recipient_id": issue.recipient_id,
                    "recipient_name": issue.recipient_name,
                    "issued_count": issue.issued_count,
                    "returned_count": issue.returned_count,
                    "issued_at": issue.issued_at.isoformat(),
                }
                for issue in self._active_issues.values()
            ],
            "history": [
                {
                    "timestamp": op.timestamp.isoformat(),
                    "operation_type": op.operation_type,
                    "apartment_id": op.apartment_id,
                    "building": op.building,
                    "floor": op.floor,
                    "apartment_number": op.apartment_number,
                    "recipient": op.recipient,
                    "quantity": op.quantity,
                    "status": op.status,
                    "details": op.details,
                }
                for op in self._history
            ],
            "next_apartment_id": self._next_apartment_id,
            "next_issue_id": self._next_issue_id,
            "next_person_id": self._next_person_id,
        }

    def load_dict(self, data: Dict[str, Any]) -> None:
        self._apartments.clear()
        self._active_issues.clear()
        self._history.clear()
        self._persons.clear()

        for item in data.get("apartments", []):
            apartment = Apartment(
                apartment_id=int(item["apartment_id"]),
                building=str(item["building"]),
                floor=int(item["floor"]),
                apartment_number=str(item["apartment_number"]),
                total_keys=int(item["total_keys"]),
                issued_keys=int(item.get("issued_keys", 0)),
                lost_keys=int(item.get("lost_keys", 0)),
            )
            self._apartments[apartment.apartment_id] = apartment

        for item in data.get("persons", []):
            person = Person(
                person_id=int(item["person_id"]),
                full_name=str(item.get("full_name", "")).strip(),
                role=str(item.get("role", "")),
                is_active=bool(item.get("is_active", True)),
            )
            if person.full_name:
                self._persons[person.person_id] = person

        for item in data.get("active_issues", []):
            recipient_name = str(item.get("recipient_name", "")).strip()
            recipient_id_raw = item.get("recipient_id")
            recipient_id = int(recipient_id_raw) if recipient_id_raw is not None else None

            if not recipient_name and recipient_id is not None and recipient_id in self._persons:
                recipient_name = self._persons[recipient_id].full_name
            if not recipient_name:
                recipient_name = str(item.get("recipient", "")).strip()

            issue = ActiveIssue(
                issue_id=int(item["issue_id"]),
                apartment_id=int(item["apartment_id"]),
                recipient_id=recipient_id,
                recipient_name=recipient_name,
                issued_count=int(item["issued_count"]),
                returned_count=int(item.get("returned_count", 0)),
                issued_at=datetime.fromisoformat(item["issued_at"]),
            )
            self._active_issues[issue.issue_id] = issue

        for item in data.get("history", []):
            apartment_id_value = item.get("apartment_id")
            apartment_id = int(apartment_id_value) if apartment_id_value is not None else None
            apartment = self._apartments.get(apartment_id) if apartment_id is not None else None

            self._history.append(
                Operation(
                    timestamp=datetime.fromisoformat(item["timestamp"]),
                    operation_type=str(item["operation_type"]),
                    apartment_id=apartment_id,
                    building=str(item.get("building", apartment.building if apartment else "-")),
                    floor=str(item.get("floor", apartment.floor if apartment else "-")),
                    apartment_number=str(item.get("apartment_number", apartment.apartment_number if apartment else "-")),
                    recipient=str(item.get("recipient", "")),
                    quantity=int(item.get("quantity", self._extract_quantity(str(item.get("details", ""))))),
                    status=str(item.get("status", self._default_status(str(item.get("operation_type", ""))))),
                    details=str(item.get("details", "")),
                )
            )

        # Миграция старого формата recipients -> persons
        for name in data.get("recipients", []):
            full_name = str(name).strip()
            if full_name and all(p.full_name != full_name for p in self._persons.values()):
                person = Person(person_id=self._next_person_id, full_name=full_name, role="", is_active=True)
                self._persons[person.person_id] = person
                self._next_person_id += 1

        computed_next_apartment_id = (max(self._apartments.keys()) + 1) if self._apartments else 1
        computed_next_issue_id = (max(self._active_issues.keys()) + 1) if self._active_issues else 1
        computed_next_person_id = (max(self._persons.keys()) + 1) if self._persons else 1

        self._next_apartment_id = max(int(data.get("next_apartment_id", 1)), computed_next_apartment_id)
        self._next_issue_id = max(int(data.get("next_issue_id", 1)), computed_next_issue_id)
        self._next_person_id = max(int(data.get("next_person_id", 1)), computed_next_person_id)

    def _get_apartment(self, apartment_id: int) -> Apartment:
        apartment = self._apartments.get(apartment_id)
        if not apartment:
            raise KeyManagerError("Квартира не найдена.")
        return apartment

    def _add_history(
        self,
        operation_type: str,
        apartment: Apartment,
        recipient: str,
        quantity: int,
        status: str,
        details: str,
    ) -> None:
        self._history.append(
            Operation(
                timestamp=datetime.now(),
                operation_type=operation_type,
                apartment_id=apartment.apartment_id,
                building=apartment.building,
                floor=str(apartment.floor),
                apartment_number=apartment.apartment_number,
                recipient=recipient,
                quantity=quantity,
                status=status,
                details=details,
            )
        )

    def _notify_change(self) -> None:
        if self._on_change:
            self._on_change()

    @staticmethod
    def _extract_quantity(details: str) -> int:
        digits = "".join(ch for ch in details if ch.isdigit())
        return int(digits) if digits else 0

    @staticmethod
    def _default_status(operation_type: str) -> str:
        return {
            "ADD_APARTMENT": "Создана",
            "ISSUE": "Выдано",
            "RETURN": "Возврат",
            "LOST": "Утеря",
        }.get(operation_type, "-")
