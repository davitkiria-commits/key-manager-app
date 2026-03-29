from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional


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
class ActiveIssue:
    issue_id: int
    apartment_id: int
    recipient: str
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
    apartment_id: int
    details: str


class KeyManagerError(Exception):
    """Base class for key manager errors."""


class KeyManager:
    def __init__(self) -> None:
        self._apartments: Dict[int, Apartment] = {}
        self._active_issues: Dict[int, ActiveIssue] = {}
        self._history: List[Operation] = []
        self._recipients: set[str] = set()

        self._next_apartment_id = 1
        self._next_issue_id = 1

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

        self._add_history(apartment.apartment_id, "ADD_APARTMENT", f"Добавлена квартира, всего ключей: {total_keys}")
        return apartment

    def issue_keys(self, apartment_id: int, recipient: str, count: int) -> ActiveIssue:
        apartment = self._get_apartment(apartment_id)

        if count <= 0:
            raise KeyManagerError("Количество выдаваемых ключей должно быть больше 0.")
        if count > apartment.available_keys:
            raise KeyManagerError(
                f"Недостаточно доступных ключей. Доступно: {apartment.available_keys}, запрошено: {count}."
            )

        recipient = recipient.strip()
        if not recipient:
            raise KeyManagerError("Укажите получателя.")

        apartment.issued_keys += count
        self._recipients.add(recipient)

        issue = ActiveIssue(
            issue_id=self._next_issue_id,
            apartment_id=apartment_id,
            recipient=recipient,
            issued_count=count,
        )
        self._active_issues[issue.issue_id] = issue
        self._next_issue_id += 1

        self._add_history(apartment_id, "ISSUE", f"Выдано {count} ключ(ей) получателю: {recipient}")
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

        self._add_history(
            apartment.apartment_id,
            "RETURN",
            f"Возвращено {count} ключ(ей) от получателя: {issue.recipient}",
        )

        if issue.active_count == 0:
            del self._active_issues[issue_id]

    def mark_lost(self, apartment_id: int, count: int) -> None:
        apartment = self._get_apartment(apartment_id)

        if count <= 0:
            raise KeyManagerError("Количество утерянных ключей должно быть больше 0.")
        if count > apartment.available_keys:
            raise KeyManagerError(
                f"Недостаточно доступных ключей для списания. Доступно: {apartment.available_keys}, запрошено: {count}."
            )

        apartment.lost_keys += count
        self._add_history(apartment_id, "LOST", f"Отмечена утеря {count} ключ(ей)")

    def get_apartments(self) -> List[Apartment]:
        return sorted(self._apartments.values(), key=lambda a: (a.building, a.floor, a.apartment_number))

    def get_apartment(self, apartment_id: int) -> Optional[Apartment]:
        return self._apartments.get(apartment_id)

    def get_active_issues(self) -> List[ActiveIssue]:
        return sorted(self._active_issues.values(), key=lambda i: i.issued_at, reverse=True)

    def get_history(self) -> List[Operation]:
        return sorted(self._history, key=lambda h: h.timestamp, reverse=True)

    def get_recipients(self) -> List[str]:
        return sorted(self._recipients)

    def _get_apartment(self, apartment_id: int) -> Apartment:
        apartment = self._apartments.get(apartment_id)
        if not apartment:
            raise KeyManagerError("Квартира не найдена.")
        return apartment

    def _add_history(self, apartment_id: int, operation_type: str, details: str) -> None:
        self._history.append(
            Operation(
                timestamp=datetime.now(),
                operation_type=operation_type,
                apartment_id=apartment_id,
                details=details,
            )
        )
