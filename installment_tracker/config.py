"""Configuration loading for the installment tracker.

All business rules live in a TOML file so the same code works for any card,
bank or business. See ``examples/config.toml`` for a documented example.
"""

from __future__ import annotations

import tomllib
from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class Config:
    """Business rules that drive every calculation."""

    #: Day of the month on which the card statement is paid (1-28).
    payment_day: int = 20
    #: Months between the purchase and its first installment
    #: (a purchase in May with offset 2 is first charged in July).
    first_installment_offset: int = 2
    #: Commission charged by the card issuer, keyed by term in months.
    #: Stored as fractions (0.025 == 2.5 %).
    commissions: dict[int, float] = field(default_factory=lambda: {1: 0.0})
    #: Allowed merchants/suppliers (used for the dropdown in the workbook).
    merchants: tuple[str, ...] = ()
    #: Primary currency symbol (what purchases are made in).
    currency_symbol: str = "$"
    #: Optional secondary currency for a converted view (e.g. local currency).
    secondary_currency_symbol: str = ""
    #: Exchange rate: 1 primary unit == ``exchange_rate`` secondary units.
    exchange_rate: float = 1.0
    #: Number of empty, formula-ready rows left in the workbook for new purchases.
    spare_rows: int = 60
    #: Number of month columns in the installment calendar.
    calendar_months: int = 24

    def commission_rate(self, term_months: int) -> float:
        """Return the commission fraction for a given term.

        Raises ``KeyError`` if the term is not configured, so a typo such as
        "5 months" never silently gets 0 % commission.
        """
        try:
            return self.commissions[term_months]
        except KeyError:
            allowed = ", ".join(str(t) for t in sorted(self.commissions))
            raise KeyError(
                f"term of {term_months} months is not in the commission table "
                f"(allowed: {allowed})"
            ) from None

    @property
    def allowed_terms(self) -> list[int]:
        return sorted(self.commissions)


def load_config(path: str | Path) -> Config:
    """Read a TOML config file and return a validated :class:`Config`."""
    raw = tomllib.loads(Path(path).read_text(encoding="utf-8"))

    commissions = {int(k): float(v) for k, v in raw.get("commissions", {}).items()}
    if not commissions:
        raise ValueError("config must define at least one [commissions] entry")
    for term, rate in commissions.items():
        if term < 1:
            raise ValueError(f"term must be >= 1 month, got {term}")
        if not 0 <= rate < 1:
            raise ValueError(
                f"commission for {term} months must be a fraction between 0 and 1 "
                f"(e.g. 0.025 for 2.5 %), got {rate}"
            )

    card = raw.get("card", {})
    payment_day = int(card.get("payment_day", 20))
    if not 1 <= payment_day <= 28:
        raise ValueError("card.payment_day must be between 1 and 28")

    currency = raw.get("currency", {})
    workbook = raw.get("workbook", {})

    return Config(
        payment_day=payment_day,
        first_installment_offset=int(card.get("first_installment_offset", 2)),
        commissions=commissions,
        merchants=tuple(raw.get("merchants", [])),
        currency_symbol=str(currency.get("symbol", "$")),
        secondary_currency_symbol=str(currency.get("secondary_symbol", "")),
        exchange_rate=float(currency.get("exchange_rate", 1.0)),
        spare_rows=int(workbook.get("spare_rows", 60)),
        calendar_months=int(workbook.get("calendar_months", 24)),
    )
