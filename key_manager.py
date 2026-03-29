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
    recipient: str = ""


class KeyManagerError(Exception):
    """Base class for key manager errors."""


class KeyManager:
    def __init__(self, on_change: Callable[[], None] | None = None) -> None:
        self._apartments: Dict[int, Apartment] = {}
        self._active_issues: Dict[int, ActiveIssue] = {}
        self._history: List[Operation] = []
        self._recipients: set[str] = set()

        self._next_apartment_id = 1
        self._next_issue_id = 1
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

        self._add_history(apartment.apartment_id, "ADD_APARTMENT", f"Добавлена квартира, всего ключей: {total_keys}")
        self._notify_change()
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

        self._add_history(apartment_id, "ISSUE", f"Выдано {count} ключ(ей) получателю: {recipient}", recipient=recipient)
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

        self._add_history(
            apartment.apartment_id,
            "RETURN",
            f"Возвращено {count} ключ(ей) от получателя: {issue.recipient}",
            recipient=issue.recipient,
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
        self._add_history(apartment_id, "LOST", f"Отмечена утеря {count} ключ(ей)")
        self._notify_change()

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
            "active_issues": [
                {
                    "issue_id": issue.issue_id,
                    "apartment_id": issue.apartment_id,
                    "recipient": issue.recipient,
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
                    "details": op.details,
                    "recipient": op.recipient,
                }
                for op in self._history
            ],
            "recipients": sorted(self._recipients),
            "next_apartment_id": self._next_apartment_id,
            "next_issue_id": self._next_issue_id,
        }

    def load_dict(self, data: Dict[str, Any]) -> None:
        self._apartments.clear()
        self._active_issues.clear()
        self._history.clear()
        self._recipients.clear()

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

        for item in data.get("active_issues", []):
            issue = ActiveIssue(
                issue_id=int(item["issue_id"]),
                apartment_id=int(item["apartment_id"]),
                recipient=str(item["recipient"]),
                issued_count=int(item["issued_count"]),
                returned_count=int(item.get("returned_count", 0)),
                issued_at=datetime.fromisoformat(item["issued_at"]),
            )
            self._active_issues[issue.issue_id] = issue

        for item in data.get("history", []):
            self._history.append(
                Operation(
                    timestamp=datetime.fromisoformat(item["timestamp"]),
                    operation_type=str(item["operation_type"]),
                    apartment_id=int(item["apartment_id"]),
                    details=str(item["details"]),
                    recipient=str(item.get("recipient", "")),
                )
            )

        self._recipients = {str(x) for x in data.get("recipients", []) if str(x).strip()}

        computed_next_apartment_id = (max(self._apartments.keys()) + 1) if self._apartments else 1
        computed_next_issue_id = (max(self._active_issues.keys()) + 1) if self._active_issues else 1

        self._next_apartment_id = max(int(data.get("next_apartment_id", 1)), computed_next_apartment_id)
        self._next_issue_id = max(int(data.get("next_issue_id", 1)), computed_next_issue_id)

    def _get_apartment(self, apartment_id: int) -> Apartment:
        apartment = self._apartments.get(apartment_id)
        if not apartment:
            raise KeyManagerError("Квартира не найдена.")
        return apartment

    def _add_history(self, apartment_id: int, operation_type: str, details: str, recipient: str = "") -> None:
        self._history.append(
            Operation(
                timestamp=datetime.now(),
                operation_type=operation_type,
                apartment_id=apartment_id,
                details=details,
                recipient=recipient,
            )
        )

    def _notify_change(self) -> None:
        if self._on_change:
            self._on_change()
