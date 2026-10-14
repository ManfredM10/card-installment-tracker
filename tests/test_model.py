from datetime import date

import pytest

from installment_tracker.model import Purchase, add_months, build_schedule


def test_add_months_crosses_year():
    assert add_months(date(2026, 11, 15), 3) == date(2027, 2, 1)


def test_commission_and_installments(config):
    s = build_schedule(Purchase("P1", date(2026, 5, 20), "Supplier A", 1664.40, 4), config)
    assert s.commission_rate == 0.025
    assert s.commission == 41.61
    assert s.total_financed == 1706.01
    assert s.monthly_installment == 416.10
    # the whole commission is charged with the first installment
    assert s.first_installment == 457.71


def test_first_and_last_month_use_offset(config):
    s = build_schedule(Purchase("P1", date(2026, 5, 20), "", 100, 4), config)
    assert s.first_month == date(2026, 7, 1)
    assert s.last_month == date(2026, 10, 1)


def test_first_month_override(config):
    s = build_schedule(Purchase("P1", date(2026, 6, 30), "", 100, 6, first_month_override=date(2026, 9, 15)), config)
    assert s.first_month == date(2026, 9, 1)
    assert s.last_month == date(2027, 2, 1)


def test_unknown_term_is_rejected(config):
    with pytest.raises(KeyError, match="5 months"):
        build_schedule(Purchase("P1", date(2026, 5, 20), "", 100, 5), config)


@pytest.mark.parametrize(
    "today, paid, status",
    [
        (date(2026, 6, 30), 0, "Pending"),   # before first month
        (date(2026, 7, 19), 0, "Pending"),   # first month, payment day not reached
        (date(2026, 7, 20), 1, "Active"),    # payment day → first installment paid
        (date(2026, 10, 19), 3, "Active"),
        (date(2026, 10, 20), 4, "Paid"),     # last installment paid
        (date(2027, 3, 1), 4, "Paid"),       # never exceeds the term
    ],
)
def test_installments_paid_inferred_from_date(config, today, paid, status):
    s = build_schedule(Purchase("P1", date(2026, 5, 20), "", 400, 4), config)
    assert s.installments_paid(today, config.payment_day) == paid
    assert s.status(today, config.payment_day) == status


def test_outstanding_balance(config):
    s = build_schedule(Purchase("P1", date(2026, 5, 20), "", 400, 4), config)
    assert s.outstanding(date(2026, 6, 1), 20) == 410.00    # nothing paid: full financed total
    assert s.outstanding(date(2026, 8, 25), 20) == 200.00   # 2 of 4 paid
    assert s.outstanding(date(2027, 1, 1), 20) == 0.0


def test_installment_for_month(config):
    s = build_schedule(Purchase("P1", date(2026, 5, 20), "", 400, 4), config)
    assert s.installment_for(date(2026, 6, 1)) == 0
    assert s.installment_for(date(2026, 7, 1)) == 110.00    # 100 + 10 commission
    assert s.installment_for(date(2026, 9, 1)) == 100.00
    assert s.installment_for(date(2026, 11, 1)) == 0


def test_next_installment_month(config):
    s = build_schedule(Purchase("P1", date(2026, 5, 20), "", 400, 4), config)
    assert s.next_installment_month(date(2026, 8, 25), 20) == date(2026, 9, 1)
    assert s.next_installment_month(date(2027, 1, 1), 20) is None
