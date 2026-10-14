from datetime import date

import pytest

from installment_tracker.config import load_config
from installment_tracker.model import read_purchases


def test_load_example_config():
    cfg = load_config("examples/config.toml")
    assert cfg.payment_day == 20
    assert cfg.commission_rate(6) == 0.03
    assert "Amazon" in cfg.merchants


def test_config_rejects_percent_instead_of_fraction(tmp_path):
    p = tmp_path / "c.toml"
    p.write_text("[commissions]\n3 = 2.5\n")
    with pytest.raises(ValueError, match="fraction"):
        load_config(p)


def test_read_example_purchases():
    purchases = read_purchases("examples/purchases.csv")
    assert len(purchases) == 15
    p7 = next(p for p in purchases if p.code == "P007")
    assert p7.first_month_override == date(2026, 9, 1)


def test_bad_row_reports_line_number(tmp_path):
    p = tmp_path / "p.csv"
    p.write_text("code,date,merchant,amount,term_months\nP1,2026-13-01,,100,3\n")
    with pytest.raises(ValueError, match="line 2"):
        read_purchases(p)
