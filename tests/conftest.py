import pytest

from installment_tracker.config import Config


@pytest.fixture
def config() -> Config:
    return Config(
        payment_day=20,
        first_installment_offset=2,
        commissions={1: 0.0, 3: 0.025, 4: 0.025, 6: 0.03},
        merchants=("Supplier A", "Supplier B"),
        currency_symbol="$",
        secondary_currency_symbol="₡",
        exchange_rate=500,
        spare_rows=3,
        calendar_months=12,
    )
