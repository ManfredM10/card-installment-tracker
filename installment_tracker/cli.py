"""Command-line entry point.

    installment-tracker build purchases.csv --config config.toml -o tracker.xlsx
    installment-tracker report purchases.csv --config config.toml [--today 2026-10-14]
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

from .config import load_config
from .excel import build_workbook
from .model import build_schedule, read_purchases


def _report(purchases, config, today: date) -> str:
    sym = config.currency_symbol
    lines = [f"{'Code':<10}{'Merchant':<12}{'Amount':>13}{'Term':>6}{'First':>10}{'Paid':>6}"
             f"{'Outstanding':>15}  Status"]
    total_out = 0.0
    for p in purchases:
        s = build_schedule(p, config)
        out = s.outstanding(today, config.payment_day)
        total_out += out
        lines.append(
            f"{p.code:<10}{(p.merchant or '-'):<12} {sym}{p.amount:>11,.2f}{p.term_months:>6}"
            f"{s.first_month.strftime('%b %Y'):>10}{s.installments_paid(today, config.payment_day):>6}"
            f" {sym}{out:>13,.2f}  {s.status(today, config.payment_day)}"
        )
    lines.append(f"\nTotal outstanding as of {today.isoformat()}: {sym}{total_out:,.2f}")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="installment-tracker",
                                     description="Track credit-card purchases financed in installments.")
    sub = parser.add_subparsers(dest="command", required=True)

    b = sub.add_parser("build", help="generate the Excel tracker")
    b.add_argument("purchases", type=Path, help="CSV with the purchases")
    b.add_argument("--config", type=Path, required=True, help="TOML file with the business rules")
    b.add_argument("-o", "--output", type=Path, default=Path("tracker.xlsx"))

    r = sub.add_parser("report", help="print a text summary (no Excel needed)")
    r.add_argument("purchases", type=Path)
    r.add_argument("--config", type=Path, required=True)
    r.add_argument("--today", type=date.fromisoformat, default=date.today(), help="override today's date (YYYY-MM-DD)")

    args = parser.parse_args(argv)
    try:
        config = load_config(args.config)
        purchases = read_purchases(args.purchases)
        # Validate every row against the rules before touching Excel.
        for p in purchases:
            build_schedule(p, config)
    except (ValueError, KeyError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    if args.command == "build":
        out = build_workbook(purchases, config, args.output)
        print(f"wrote {out} ({len(purchases)} purchases)")
    else:
        print(_report(purchases, config, args.today))
    return 0


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
