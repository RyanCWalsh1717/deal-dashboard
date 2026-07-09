"""Extraction of cash-account balances (escrows, reserves, operating cash) and
loan-statement data from two real, confirmed formats (validated 2026-07-09
against actual Revolution Labs documents):

1. Berkadia loan servicer statements (PDF) — one per tranche, e.g. "Revolution
   Labs - Note A1". Fixed two-column layout; only the left column (balance
   info) is parsed. Confirmed fields: Principal Balance, Interest Rate (the
   real all-in rate — NOT the same as the "rate" derived in debt.py from the
   distribution workbook, which is actually just the spread over SOFR), Tax
   Escrow Balance, Insurance Escrow Balance, Reserve Balance.
2. A Yardi trial balance export (.xlsx) — account code / label / Forward /
   Debit / Credit / Ending Balance columns, with an entity header row
   ("Property = <code> ...") marking the start of that entity's account
   block. A single export can cover more than one property, so this is
   entity-block-aware rather than a blind whole-file scan — extraction is
   filtered to the requested property_code's block only.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Union

from pipeline.models import CashAccountBalance, LoanStatement

_HEADER_RE = re.compile(
    r"Property:\s*(?P<label>.+?)\s+Loan No:\s*(?P<loan_no>\S+)\s+Interest Rate:\s*(?P<rate>[\d.]+)"
)
_AS_OF_RE = re.compile(r"AS OF\s+(\d{2}/\d{2}/\d{4})")


def _money(text: str, label: str) -> Optional[float]:
    match = re.search(re.escape(label) + r"\s+\$?\s*([\d,]+\.\d{2})", text)
    return float(match.group(1).replace(",", "")) if match else None


def parse_loan_statement(path: Union[str, Path]) -> Optional[LoanStatement]:
    import pdfplumber

    with pdfplumber.open(path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    header = _HEADER_RE.search(text)
    if not header:
        return None

    property_label = header.group("label").strip()
    tranche_name = property_label.split(" - ")[-1].strip() if " - " in property_label else property_label

    as_of_match = _AS_OF_RE.search(text)
    as_of = datetime.strptime(as_of_match.group(1), "%m/%d/%Y").date() if as_of_match else None

    return LoanStatement(
        tranche_name=tranche_name,
        loan_number=header.group("loan_no"),
        interest_rate=float(header.group("rate")) / 100.0,
        as_of=as_of,
        principal_balance=_money(text, "Principal Balance"),
        interest_paid_ytd=_money(text, "Interest Paid YTD"),
        tax_escrow_balance=_money(text, "Tax Escrow Balance"),
        insurance_escrow_balance=_money(text, "Insurance Escrow Balance"),
        reserve_balance=_money(text, "Reserve Balance"),
        total_payment_due=_money(text, "Total Payment Due"),
    )


def loan_statement_cash_accounts(stmt: LoanStatement) -> List[CashAccountBalance]:
    """The escrow/reserve lines of a loan statement, as Cash-tab boxes."""
    boxes = []
    for label, value in (
        (f"Tax Escrow ({stmt.tranche_name})", stmt.tax_escrow_balance),
        (f"Insurance Escrow ({stmt.tranche_name})", stmt.insurance_escrow_balance),
        (f"Reserve ({stmt.tranche_name})", stmt.reserve_balance),
    ):
        if value is not None:
            boxes.append(
                CashAccountBalance(label=label, balance=value, source="loan_statement", as_of=stmt.as_of)
            )
    return boxes


_ENTITY_HEADER_RE = re.compile(r"Property\s*=\s*(\S+)")
_CASH_LABEL_KEYWORDS = ("cash - operating", "cash - development", "restricted cash", "escrow", "reserve")


def parse_trial_balance_cash_accounts(
    path: Union[str, Path], yardi_codes: Optional[List[str]] = None
) -> List[CashAccountBalance]:
    """Scans a Yardi trial balance export for cash/escrow/reserve accounts,
    restricted to entity block(s) matching one of `yardi_codes` (a TB export
    can cover more than one property, and the dashboard's own property_code
    is just an internal folder name that won't match Yardi's real codes — see
    PropertyConfig.yardi_codes). If `yardi_codes` is empty/None, or no
    matching block is found, falls back to scanning the whole sheet (better
    than silently returning nothing, but callers should prefer passing it)."""
    import openpyxl

    yardi_codes = yardi_codes or []
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]

    as_of = None
    for row in ws.iter_rows(min_row=1, max_row=6, max_col=1):
        v = row[0].value
        if isinstance(v, str) and v.strip().lower().startswith("period"):
            period_text = v.split("=", 1)[1].strip() if "=" in v else v.strip()
            try:
                as_of = datetime.strptime(period_text, "%B %Y").date()
            except ValueError:
                pass

    # Column layout: 1=code, 2=label, 3=forward, 4=debit, 5=credit, 6=ending balance.
    results: List[CashAccountBalance] = []
    in_target_block = not yardi_codes
    for r in range(1, ws.max_row + 1):
        col1 = ws.cell(row=r, column=1).value
        if isinstance(col1, str) and col1.strip().lower().startswith("property"):
            entity_match = _ENTITY_HEADER_RE.search(col1)
            block_code = entity_match.group(1) if entity_match else col1
            in_target_block = not yardi_codes or any(code in block_code for code in yardi_codes)
            continue

        if not in_target_block:
            continue

        label = ws.cell(row=r, column=2).value
        ending = ws.cell(row=r, column=6).value
        if not isinstance(label, str) or not isinstance(ending, (int, float)):
            continue
        if not any(kw in label.lower() for kw in _CASH_LABEL_KEYWORDS):
            continue

        results.append(
            CashAccountBalance(
                label=label.strip(),
                balance=float(ending),
                account_code=str(col1) if col1 is not None else "",
                source="trial_balance",
                as_of=as_of,
            )
        )

    return results
