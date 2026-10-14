"""Pure business logic: how a financed purchase turns into installments.

Everything here is plain Python with no spreadsheet dependency, so it can be
unit-tested in isolation. The Excel builder mirrors these rules with formulas.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import date
from pathlib import Path

from .config import Config


def add_months(d: date, months: int) -> date:
    """Return the first day of the month ``months`` after ``d``'s month."""
    total = d.year * 12 + (d.month - 1) + months
    return date(total // 12, total % 12 + 1, 1)


def months_between(start: date, end: date) -> int:
    """Whole months from ``start``'s month to ``end``'s month (can be negative)."""
    return (end.year - start.year) * 12 + (end.month - start.month)


@dataclass(frozen=True)
class Purchase:
    """A single purchase financed through the card."""

    code: str
    purchase_date: date
    merchant: str
    amount: float
    term_months: int
    #: Optional override: first day of the month the bank actually charged the
    #: first installment, when it differs from the configured offset.
    first_month_override: date | None = None

    def __post_init__(self) -> None:
        if self.amount <= 0:
            raise ValueError(f"{self.code}: amount must be positive")
        if self.term_months < 1:
            raise ValueError(f"{self.code}: term must be at least 1 month")


@dataclass(frozen=True)
class Schedule:
    """Everything derived from a purchase under a given :class:`Config`."""

    purchase: Purchase
    commission_rate: float
    commission: float
    total_financed: float
    monthly_installment: float
    first_installment: float
    first_month: date
    last_month: date

    def installment_for(self, month: date) -> float:
        """Amount due in ``month`` (first day of month), or 0 if none."""
        m = date(month.year, month.month, 1)
        if m == self.first_month:
            return self.first_installment
        if self.first_month < m <= self.last_month:
            return self.monthly_installment
        return 0.0

    def installments_paid(self, today: date, payment_day: int) -> int:
        """How many installments are already paid, inferred from the date.

        An installment is considered paid once the card's payment day of that
        month has passed. No manual marking needed.
        """
        elapsed = months_between(self.first_month, today)
        if today.day >= payment_day:
            elapsed += 1
        return max(0, min(self.purchase.term_months, elapsed))

    def status(self, today: date, payment_day: int) -> str:
        paid = self.installments_paid(today, payment_day)
        if paid == 0:
            return "Pending"
        if paid >= self.purchase.term_months:
            return "Paid"
        return "Active"

    def outstanding(self, today: date, payment_day: int) -> float:
        paid = self.installments_paid(today, payment_day)
        if paid == 0:
            return round(self.total_financed, 2)
        remaining = self.purchase.term_months - paid
        return round(remaining * self.monthly_installment, 2)

    def next_installment_month(self, today: date, payment_day: int) -> date | None:
        paid = self.installments_paid(today, payment_day)
        if paid >= self.purchase.term_months:
            return None
        return add_months(self.first_month, paid)


def build_schedule(purchase: Purchase, config: Config) -> Schedule:
    """Apply the configured rules to a purchase."""
    rate = config.commission_rate(purchase.term_months)
    commission = round(purchase.amount * rate, 2)
    monthly = round(purchase.amount / purchase.term_months, 2)
    if purchase.first_month_override is not None:
        first = date(purchase.first_month_override.year, purchase.first_month_override.month, 1)
    else:
        first = add_months(purchase.purchase_date, config.first_installment_offset)
    last = add_months(first, purchase.term_months - 1)
    return Schedule(
        purchase=purchase,
        commission_rate=rate,
        commission=commission,
        total_financed=round(purchase.amount + commission, 2),
        monthly_installment=monthly,
        # The issuer charges the whole commission together with the first installment.
        first_installment=round(monthly + commission, 2),
        first_month=first,
        last_month=last,
    )


CSV_COLUMNS = ("code", "date", "merchant", "amount", "term_months", "first_month_override")


def read_purchases(path: str | Path) -> list[Purchase]:
    """Load purchases from a CSV with the columns in :data:`CSV_COLUMNS`.

    Dates are ISO (``YYYY-MM-DD``). ``merchant`` and ``first_month_override``
    may be empty.
    """
    purchases: list[Purchase] = []
    with Path(path).open(newline="", encoding="utf-8") as fh:
        reader = csv.DictReader(fh)
        missing = set(CSV_COLUMNS[:5]) - set(reader.fieldnames or [])
        if missing:
            raise ValueError(f"CSV is missing columns: {', '.join(sorted(missing))}")
        for line, row in enumerate(reader, start=2):
            try:
                override = row.get("first_month_override") or ""
                purchases.append(
                    Purchase(
                        code=row["code"].strip(),
                        purchase_date=date.fromisoformat(row["date"].strip()),
                        merchant=(row.get("merchant") or "").strip(),
                        amount=float(row["amount"]),
                        term_months=int(row["term_months"]),
                        first_month_override=date.fromisoformat(override.strip()) if override.strip() else None,
                    )
                )
            except (KeyError, ValueError) as exc:
                raise ValueError(f"{path}, line {line}: {exc}") from None
    return purchases
