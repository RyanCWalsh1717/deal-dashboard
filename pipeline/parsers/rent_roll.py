"""Parses a Yardi "Tenancy Schedule II" rent roll export into a RentRollResult.

Layout observed (`TenancyScheduleII07_09_2026.xlsx`, sheet "Report1"): a
multi-row-per-lease block. The FIRST row of each block carries the property
name in column 1 (the reliable anchor for "this is a new lease") plus the
core lease fields. Every other row in that block (blank column 1) up to the
next primary row carries one or both of: a "Charge Type" line (col 18/19 —
"Rent" or "CAM"; a lease can have several CAM lines, e.g. RE-Tax + Operating
components, which the export never distinguishes beyond "not Rent") and/or a
dated future rent-escalation step (col 25 start date, col 31/32 annual
amount/PSF) — plus occasionally a second physical unit on a multi-suite
lease (e.g. an office suite plus a penthouse), whose own area is already
folded into the primary row's `lease_area` (col 13), so its row doesn't need
separate parsing.

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
from typing import List, Optional

from openpyxl.worksheet.worksheet import Worksheet

from pipeline.models import ChargeLine, RentRollLine, RentRollResult, RentStep

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


def _is_primary_row(ws: Worksheet, r: int) -> bool:
    property_cell = ws.cell(row=r, column=1).value
    if not isinstance(property_cell, str) or not property_cell.strip():
        return False
    tenant_raw = ws.cell(row=r, column=7).value
    tenant_name = tenant_raw.strip() if isinstance(tenant_raw, str) else ""
    is_vacant = tenant_name.upper() == "VACANT"
    # Column 1 is non-empty on the title/property-header/column-header rows
    # too (rows 1-3) — a real lease row always has either a lease-from date
    # or the literal "VACANT" marker, which those rows never do.
    lease_from_cell = ws.cell(row=r, column=9).value
    return is_vacant or isinstance(lease_from_cell, datetime)


def _parse_block_charges_and_steps(ws: Worksheet, start: int, end: int):
    """Scans rows [start, end) — a lease's primary row plus every row up to
    (not including) the next lease's primary row — for charge-type lines and
    rent-escalation steps. The two checks are independent per row (rows exist
    with both populated at once), so neither classification excludes the
    other."""
    charges: List[ChargeLine] = []
    rent_steps: List[RentStep] = []
    for r in range(start, end):
        charge_type = ws.cell(row=r, column=18).value
        charge_amt = ws.cell(row=r, column=19).value
        if isinstance(charge_type, str) and charge_type.strip() and isinstance(charge_amt, (int, float)):
            charges.append(ChargeLine(charge_type=charge_type.strip(), annual_amount=float(charge_amt)))

        step_date = ws.cell(row=r, column=25).value
        step_rent = ws.cell(row=r, column=31).value
        if isinstance(step_date, datetime) and isinstance(step_rent, (int, float)):
            rent_steps.append(
                RentStep(
                    effective_date=step_date.date(),
                    annual_rent=float(step_rent),
                    annual_rent_psf=_to_float(ws.cell(row=r, column=32).value),
                )
            )
    rent_steps.sort(key=lambda s: s.effective_date)
    return charges, rent_steps


def parse_rent_roll(ws: Worksheet, property_code: str) -> RentRollResult:
    as_of = _parse_as_of(ws)
    primary_rows = [r for r in range(1, ws.max_row + 1) if _is_primary_row(ws, r)]

    lines = []
    for idx, r in enumerate(primary_rows):
        block_end = primary_rows[idx + 1] if idx + 1 < len(primary_rows) else ws.max_row + 1

        tenant_raw = ws.cell(row=r, column=7).value
        tenant_name = tenant_raw.strip() if isinstance(tenant_raw, str) else ""
        is_vacant = tenant_name.upper() == "VACANT"

        charges, rent_steps = _parse_block_charges_and_steps(ws, r, block_end)

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
                loc_amount=_to_float(ws.cell(row=r, column=17).value),
                charges=charges,
                rent_steps=rent_steps,
            )
        )

    return RentRollResult(property_code=property_code, as_of=as_of, lines=lines)
