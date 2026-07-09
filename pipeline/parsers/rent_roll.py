"""Parses a Yardi "Tenancy Schedule II" rent roll export into a RentRollResult.

Layout observed (`TenancyScheduleII07_09_2026.xlsx`, sheet "Report1"): a
multi-row-per-lease block. The FIRST row of each block carries the property
name in column 1 (the reliable anchor for "this is a new lease") plus the
core lease fields; any following rows in the same block (blank column 1) are
charge-type detail (Rent, CAM, ...) or a second unit on a multi-suite lease —
not needed for a basic rent roll view, so only primary rows are parsed here.

Column mapping below is empirical, not derived from the printed header row —
the header row's "Unit Type" column is never actually populated in the data,
which shifts data one column left of where the header text would suggest
(header says "Lease" in column 7 / "Customer" in column 8; the real tenant
name is in column 7). Confirmed directly against the real file rather than
trusted from the header labels.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Optional

from openpyxl.worksheet.worksheet import Worksheet

from pipeline.models import RentRollLine, RentRollResult

_AS_OF_RE = re.compile(r"As of Date:\s*(\d{1,2}/\d{1,2}/\d{4})")


def _parse_as_of(ws: Worksheet) -> Optional[date]:
    for row in ws.iter_rows(min_row=1, max_row=4, max_col=1):
        v = row[0].value
        if isinstance(v, str):
            match = _AS_OF_RE.search(v)
            if match:
                return datetime.strptime(match.group(1), "%m/%d/%Y").date()
    return None


def _to_date(v) -> Optional[date]:
    return v.date() if isinstance(v, datetime) else None


def _to_float(v) -> Optional[float]:
    return float(v) if isinstance(v, (int, float)) else None


def parse_rent_roll(ws: Worksheet, property_code: str) -> RentRollResult:
    as_of = _parse_as_of(ws)
    lines = []

    for r in range(1, ws.max_row + 1):
        property_cell = ws.cell(row=r, column=1).value
        if not isinstance(property_cell, str) or not property_cell.strip():
            continue  # not a primary lease row

        tenant_raw = ws.cell(row=r, column=7).value
        tenant_name = tenant_raw.strip() if isinstance(tenant_raw, str) else ""
        is_vacant = tenant_name.upper() == "VACANT"

        # Column 1 is non-empty on the title/property-header/column-header rows
        # too (rows 1-3) — a real lease row always has either a lease-from
        # date or the literal "VACANT" marker, which those rows never do.
        lease_from_cell = ws.cell(row=r, column=9).value
        if not is_vacant and not isinstance(lease_from_cell, datetime):
            continue

        lines.append(
            RentRollLine(
                building=str(ws.cell(row=r, column=2).value or "").strip(),
                floor=str(ws.cell(row=r, column=3).value or "").strip(),
                unit_code=str(ws.cell(row=r, column=4).value or "").strip(),
                unit_area=_to_float(ws.cell(row=r, column=6).value),
                tenant_name="" if is_vacant else tenant_name,
                lease_from=_to_date(ws.cell(row=r, column=9).value),
                lease_to=_to_date(ws.cell(row=r, column=10).value),
                term_months=_to_float(ws.cell(row=r, column=11).value),
                lease_area=_to_float(ws.cell(row=r, column=13).value),
                annual_rent=_to_float(ws.cell(row=r, column=14).value),
                annual_rent_psf=_to_float(ws.cell(row=r, column=15).value),
                lease_type=str(ws.cell(row=r, column=16).value or "").strip(),
                is_vacant=is_vacant,
            )
        )

    return RentRollResult(property_code=property_code, as_of=as_of, lines=lines)
