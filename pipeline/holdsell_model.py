"""Hold/Sell pro forma: assumption storage + Excel export.

The calculation logic (rollover schedule, debt amortization, IRR) lives
entirely in the exported workbook's own formulas — deliberately not
reimplemented here, so there is exactly one place the math can be wrong,
not two that could quietly disagree. This module only (1) persists the
assumption *values* Ryan edits in the app, and (2) writes those values,
plus whatever real data the app already has loaded, into a freshly-built
copy of that same workbook.

Sheet layout and formulas below are the same ones built and manually
audited on 2026-08-18 (this sandbox has no LibreOffice, so recalc.py can't
run here — see that session's notes on the manual audit approach used
instead: full formula dump + defined-names cross-reference).
"""

from __future__ import annotations

import math
from pathlib import Path
from typing import Dict, List, Optional

import openpyxl
import yaml
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.workbook import Workbook

from pipeline.models import DistributionWorkbookResult, EntityTrialBalance, LoanStatement, RentRollResult
from pipeline.property_config import PropertyConfig

N_YEARS = 10  # projection years modeled; Hold Period (input) gates which of these actually count
YEAR_COLS = list(range(3, 3 + N_YEARS))  # columns C..L

# (key, label, section, kind) — kind is one of "years" | "pct" | "months" | "dollar" | "dollar_psf".
# Drives both the Streamlit form (grouped by section) and the Excel cell writes below — a single
# source of truth so the two never drift out of sync with each other.
ASSUMPTION_FIELDS = [
    ("hold_period", "Hold Period (years)", "Hold Period", "years"),
    ("rent_growth", "Rent Growth — Market Rent, annual %", "Inflation / Growth", "pct"),
    ("opex_growth", "Opex Growth, annual %", "Inflation / Growth", "pct"),
    ("tilc_growth", "TI / LC Growth, annual %", "Inflation / Growth", "pct"),
    ("capital_growth", "Capital Growth, annual %", "Inflation / Growth", "pct"),
    ("vacancy_credit", "General Vacancy & Credit Loss, % of collected rent", "Vacancy Credit", "pct"),
    ("renewal_downtime", "Downtime (months) — Renewal", "Market Leasing Assumptions", "months"),
    ("new_downtime", "Downtime (months) — New Lease", "Market Leasing Assumptions", "months"),
    ("renewal_free_rent", "Free Rent (months) — Renewal", "Market Leasing Assumptions", "months"),
    ("new_free_rent", "Free Rent (months) — New Lease", "Market Leasing Assumptions", "months"),
    ("renewal_ti", "TI ($/SF) — Renewal", "Market Leasing Assumptions", "dollar_psf"),
    ("new_ti", "TI ($/SF) — New Lease", "Market Leasing Assumptions", "dollar_psf"),
    ("renewal_lc", "LC (% of Year-1 rent) — Renewal", "Market Leasing Assumptions", "pct"),
    ("new_lc", "LC (% of Year-1 rent) — New Lease", "Market Leasing Assumptions", "pct"),
    ("renewal_prob", "Renewal Probability, %", "Market Leasing Assumptions", "pct"),
    ("market_rent_y1", "Market Rent, Year 1 ($/SF/Yr)", "Market Leasing Assumptions", "dollar_psf"),
    ("recovery_pct", "Recovery %, of Opex", "Market Leasing Assumptions", "pct"),
    ("cap_rate", "Exit Cap Rate, %", "Exit Assumptions", "pct"),
    ("cost_of_sale", "Cost of Sale, % of gross sale price", "Exit Assumptions", "pct"),
    ("refi_year", "Refinance Year (1-10; 0 = none modeled)", "Refinance Assumptions", "years"),
    ("refi_amount", "New Loan Amount ($)", "Refinance Assumptions", "dollar"),
    ("refi_rate", "New Loan Rate, %", "Refinance Assumptions", "pct"),
    ("refi_amort_years", "New Amortization (years)", "Refinance Assumptions", "years"),
    ("refi_io_years", "Interest-Only Period (years)", "Refinance Assumptions", "years"),
]

DEFAULT_ASSUMPTIONS: Dict[str, Optional[float]] = {key: None for key, *_ in ASSUMPTION_FIELDS}

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1A5C22")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=14, color="1A5C22")
SECTION_FONT = Font(name=FONT_NAME, bold=True, size=11, color="1A5C22")
NOTE_FONT = Font(name=FONT_NAME, italic=True, size=9, color="666666")
INPUT_FONT = Font(name=FONT_NAME, color="0000FF")
FORMULA_FONT = Font(name=FONT_NAME, color="000000")
BOLD_FONT = Font(name=FONT_NAME, bold=True)
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
SECTION_FILL = PatternFill("solid", fgColor="E8F0E9")
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY_FMT = '$#,##0;($#,##0);"-"'
MONEY_PSF_FMT = '$#,##0.00;($#,##0.00);"-"'
PCT_FMT = "0.0%"
PCT1_FMT = "0.00%"
SF_FMT = "#,##0"
MULT_FMT = "0.00x"


def _assumptions_path(cfg: PropertyConfig, data_dir: str = "data") -> Path:
    return Path(data_dir) / cfg.property_code / "holdsell_assumptions.yaml"


def load_assumptions(cfg: PropertyConfig, data_dir: str = "data") -> Dict[str, Optional[float]]:
    path = _assumptions_path(cfg, data_dir)
    values = dict(DEFAULT_ASSUMPTIONS)
    if path.exists():
        with open(path, "r") as f:
            saved = yaml.safe_load(f) or {}
        for key in DEFAULT_ASSUMPTIONS:
            if key in saved:
                values[key] = saved[key]
    return values


def save_assumptions(cfg: PropertyConfig, assumptions: Dict[str, Optional[float]], data_dir: str = "data") -> None:
    path = _assumptions_path(cfg, data_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w") as f:
        yaml.safe_dump(assumptions, f, default_flow_style=False, sort_keys=False)


def _style_header_row(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def _section_header(ws, row, text, span=2):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    for c in range(1, span + 1):
        ws.cell(row=row, column=c).fill = SECTION_FILL
    return row + 1


def _autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _input_cell(ws, row, col, value=None, fmt=None, comment=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    if comment:
        cell.comment = Comment(comment, "Deal Dashboard")
    return cell


def _real_cell(ws, row, col, value, fmt=None, source_comment=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = FORMULA_FONT
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    if source_comment:
        cell.comment = Comment(source_comment, "Deal Dashboard")
    return cell


def _formula_cell(ws, row, col, formula, fmt=None, bold=False):
    cell = ws.cell(row=row, column=col, value=formula)
    cell.font = BOLD_FONT if bold else FORMULA_FONT
    cell.border = BORDER
    if fmt:
        cell.number_format = fmt
    return cell


def _year_header_row(ws, row, n_years=N_YEARS):
    ws.cell(row=row, column=1, value="Year")
    ws.cell(row=row, column=2, value="Year 0\n(Current)")
    for i, col in enumerate(YEAR_COLS, start=1):
        ws.cell(row=row, column=col, value=f"Year {i}")
    _style_header_row(ws, row, 2 + n_years, start_col=1)
    ws.row_dimensions[row].height = 26


def _col(c):
    return get_column_letter(c)


def build_workbook(
    cfg: PropertyConfig,
    result: Optional[DistributionWorkbookResult],
    rent_roll: Optional[RentRollResult],
    loan_statements: List[LoanStatement],
    entity_trial_balances: List[EntityTrialBalance],
    assumptions: Dict[str, Optional[float]],
) -> Workbook:
    if not result or not result.annual_budget_summary or not rent_roll:
        raise ValueError("build_workbook() needs a parsed distribution workbook (with annual budget summary) and rent roll.")

    a = dict(DEFAULT_ASSUMPTIONS)
    a.update({k: v for k, v in assumptions.items() if k in DEFAULT_ASSUMPTIONS})

    annual = {l.account_code: l for l in result.annual_budget_summary.lines}
    as_of = rent_roll.as_of
    ytd_through = result.annual_budget_summary.period
    months_ytd = ytd_through.month if ytd_through else 12
    tb = entity_trial_balances[0] if entity_trial_balances else None

    def rollover_year(line):
        if line.is_vacant or not line.lease_to or not as_of:
            return 1
        days = (line.lease_to - as_of).days
        years = days / 365.25
        return max(1, math.ceil(years))

    out = openpyxl.Workbook()
    out.remove(out.active)
    names: Dict[str, str] = {}

    def define(name, sheet, col, row):
        names[name] = f"'{sheet}'!${_col(col)}${row}"

    # =======================================================================
    # Sheet 1: Assumptions
    # =======================================================================
    wsA = out.create_sheet("Assumptions")
    wsA.sheet_view.showGridLines = False
    _autosize(wsA, [42, 20, 20, 55])

    wsA["A1"] = f"{cfg.display()} — Hold/Sell Model Assumptions"
    wsA["A1"].font = TITLE_FONT
    wsA["A2"] = (
        "Legend: yellow cells are your inputs, saved from the Deal Dashboard's Hold/Sell "
        "Assumptions page. Everything else is real data from the app's parsed source files, "
        "cited in each cell's comment."
    )
    wsA["A2"].font = NOTE_FONT
    wsA.merge_cells("A2:D2")
    wsA.row_dimensions[2].height = 28

    r = 4
    r = _section_header(wsA, r, "Hold Period", span=4)
    wsA.cell(row=r, column=1, value="Hold Period (years)")
    _input_cell(wsA, r, 2, value=a["hold_period"], fmt="0", comment="How many of the 10 projection years count toward the exit/IRR calc.")
    define("HoldPeriod", "Assumptions", 2, r)
    r += 2

    r = _section_header(wsA, r, "Inflation / Growth Assumptions", span=4)
    for label, name, key in [
        ("Rent Growth — Market Rent, annual %", "RentGrowth", "rent_growth"),
        ("Opex Growth, annual %", "OpexGrowth", "opex_growth"),
        ("TI / LC Growth, annual %", "TILCGrowth", "tilc_growth"),
        ("Capital Growth, annual %", "CapitalGrowth", "capital_growth"),
    ]:
        wsA.cell(row=r, column=1, value=label)
        _input_cell(wsA, r, 2, value=a[key], fmt=PCT_FMT)
        define(name, "Assumptions", 2, r)
        r += 1
    r += 1

    r = _section_header(wsA, r, "Vacancy Credit", span=4)
    wsA.cell(row=r, column=1, value="General Vacancy & Credit Loss, % of collected rent")
    _input_cell(wsA, r, 2, value=a["vacancy_credit"], fmt=PCT_FMT, comment=(
        "Frictional vacancy/bad-debt allowance applied on TOP of the explicit suite-level "
        "downtime already modeled on the Rollover sheet — this is the general cushion, not the "
        "known rollover gaps."
    ))
    define("VacancyCredit", "Assumptions", 2, r)
    r += 2

    r = _section_header(wsA, r, "Market Leasing Assumptions", span=4)
    wsA.cell(row=r, column=1, value="Applied at every suite's rollover — blended by Renewal Probability below.").font = NOTE_FONT
    r += 1
    hdr_row = r
    for i, h in enumerate(["", "Renewal", "New Lease (post-downtime)"], start=1):
        wsA.cell(row=hdr_row, column=i, value=h)
    _style_header_row(wsA, hdr_row, 3)
    r += 1
    for label, key_suffix, ren_key, new_key in [
        ("Downtime (months)", "Downtime", "renewal_downtime", "new_downtime"),
        ("Free Rent (months)", "FreeRent", "renewal_free_rent", "new_free_rent"),
        ("TI ($/SF)", "TI", "renewal_ti", "new_ti"),
        ("LC (% of Year-1 rent)", "LC", "renewal_lc", "new_lc"),
    ]:
        wsA.cell(row=r, column=1, value=label)
        fmt = PCT_FMT if key_suffix == "LC" else (MONEY_PSF_FMT if key_suffix == "TI" else "0.0")
        _input_cell(wsA, r, 2, value=a[ren_key], fmt=fmt)
        _input_cell(wsA, r, 3, value=a[new_key], fmt=fmt)
        define(f"Renewal{key_suffix}", "Assumptions", 2, r)
        define(f"New{key_suffix}", "Assumptions", 3, r)
        r += 1
    wsA.cell(row=r, column=1, value="Renewal Probability, %")
    _input_cell(wsA, r, 2, value=a["renewal_prob"], fmt=PCT_FMT, comment="Probability an expiring/vacant suite renews vs. re-leases as a new lease (post-downtime).")
    define("RenewalProb", "Assumptions", 2, r)
    r += 1
    wsA.cell(row=r, column=1, value="Market Rent, Year 1 ($/SF/Yr)")
    _input_cell(wsA, r, 2, value=a["market_rent_y1"], fmt=MONEY_PSF_FMT, comment=(
        "Market rent for a suite rolling over in Year 1. Grows by Rent Growth each year after "
        "that — applies to both Renewal and New Lease outcomes; both reset to market, they "
        "differ in downtime/free rent/TI/LC."
    ))
    define("MarketRentY1", "Assumptions", 2, r)
    r += 1
    wsA.cell(row=r, column=1, value="Recovery %, of Opex")
    _input_cell(wsA, r, 2, value=a["recovery_pct"], fmt=PCT_FMT, comment=(
        "Blended recovery structure — % of operating expenses recouped from tenants as "
        "reimbursement revenue. Simplification: applied uniformly rather than lease-by-lease; "
        "scaled down in rollover years by the same rent-collection factor as the suite's own "
        "downtime (see Cash Flow Projection)."
    ))
    define("RecoveryPct", "Assumptions", 2, r)
    r += 2

    r = _section_header(wsA, r, "Exit Assumptions", span=4)
    wsA.cell(row=r, column=1, value="Exit Cap Rate, %")
    _input_cell(wsA, r, 2, value=a["cap_rate"], fmt=PCT_FMT, comment="Applied to the exit year's own (trailing) NOI — not forward NOI.")
    define("CapRate", "Assumptions", 2, r)
    r += 1
    wsA.cell(row=r, column=1, value="Cost of Sale, % of gross sale price")
    _input_cell(wsA, r, 2, value=a["cost_of_sale"], fmt=PCT_FMT)
    define("CostOfSale", "Assumptions", 2, r)
    r += 2

    r = _section_header(wsA, r, "Refinance Assumptions", span=4)
    wsA.cell(row=r, column=1, value="Refinance Year (1-10; 0 or blank = no refinance modeled)")
    _input_cell(wsA, r, 2, value=a["refi_year"], fmt="0")
    define("RefiYear", "Assumptions", 2, r)
    r += 1
    wsA.cell(row=r, column=1, value="New Loan Amount ($)")
    _input_cell(wsA, r, 2, value=a["refi_amount"], fmt=MONEY_FMT)
    define("RefiAmount", "Assumptions", 2, r)
    r += 1
    wsA.cell(row=r, column=1, value="New Loan Rate, %")
    _input_cell(wsA, r, 2, value=a["refi_rate"], fmt=PCT_FMT)
    define("RefiRate", "Assumptions", 2, r)
    r += 1
    wsA.cell(row=r, column=1, value="New Amortization (years)")
    _input_cell(wsA, r, 2, value=a["refi_amort_years"], fmt="0")
    define("RefiAmortYears", "Assumptions", 2, r)
    r += 1
    wsA.cell(row=r, column=1, value="Interest-Only Period (years)")
    _input_cell(wsA, r, 2, value=a["refi_io_years"], fmt="0")
    define("RefiIOYears", "Assumptions", 2, r)
    r += 2

    r = _section_header(wsA, r, "Starting Position — Real Data", span=4)
    hdr_row = r
    for i, h in enumerate(["Existing Debt Tranche", "Balance", "Rate"], start=1):
        wsA.cell(row=hdr_row, column=i, value=h)
    _style_header_row(wsA, hdr_row, 3)
    r += 1
    tranche_first_row = r
    for stmt in loan_statements:
        wsA.cell(row=r, column=1, value=stmt.tranche_name).border = BORDER
        _real_cell(wsA, r, 2, stmt.principal_balance, MONEY_FMT, "Source: loan servicer statement, currently loaded period.")
        _real_cell(wsA, r, 3, stmt.interest_rate, PCT1_FMT, "Source: loan servicer statement (real all-in rate).")
        r += 1
    if not loan_statements:
        wsA.cell(row=r, column=1, value="(no loan statements loaded)").font = NOTE_FONT
        r += 1
    tranche_last_row = r - 1
    wsA.cell(row=r, column=1, value="Total / Blended").font = BOLD_FONT
    if loan_statements:
        _formula_cell(wsA, r, 2, f"=SUM(B{tranche_first_row}:B{tranche_last_row})", MONEY_FMT, bold=True)
        _formula_cell(
            wsA, r, 3,
            f"=SUMPRODUCT(B{tranche_first_row}:B{tranche_last_row},C{tranche_first_row}:C{tranche_last_row})/B{r}",
            PCT1_FMT, bold=True,
        )
    else:
        _formula_cell(wsA, r, 2, "=0", MONEY_FMT, bold=True)
        _formula_cell(wsA, r, 3, "=0", PCT1_FMT, bold=True)
    define("ExistingDebtBalance", "Assumptions", 2, r)
    define("ExistingBlendedRate", "Assumptions", 3, r)
    r += 2

    debt_service_actual = annual["debt_service"].actual_value if annual.get("debt_service") else None
    wsA.cell(row=r, column=1, value="Existing Annual Debt Service (annualized run-rate, real)")
    _real_cell(
        wsA, r, 2, (debt_service_actual or 0) / months_ytd * 12 if debt_service_actual else 0, MONEY_FMT,
        f"Source: distribution workbook, YTD actual Debt Service through {ytd_through}, annualized "
        f"(x12/{months_ytd}). Modeled as interest-only through refinance — real debt service ties "
        "closely to pure interest on the balance above, supporting this simplification."
    )
    define("ExistingDebtService", "Assumptions", 2, r)
    r += 2

    contributions = tb.contributions if tb else None
    distributions = tb.distributions if tb else None
    wsA.cell(row=r, column=1, value=f"Contributions to Date ({tb.entity_code if tb else 'n/a'}, real)")
    _real_cell(wsA, r, 2, contributions or 0, MONEY_FMT, "Source: trial balance, currently loaded period.")
    contrib_row = r
    r += 1
    wsA.cell(row=r, column=1, value=f"Distributions to Date ({tb.entity_code if tb else 'n/a'}, real)")
    _real_cell(wsA, r, 2, distributions or 0, MONEY_FMT, "Source: trial balance, currently loaded period.")
    distrib_row = r
    r += 1
    wsA.cell(row=r, column=1, value="Starting Net Invested Equity (Contributions − Distributions)")
    _formula_cell(wsA, r, 2, f"=B{contrib_row}-B{distrib_row}", MONEY_FMT, bold=True)
    wsA.cell(row=r, column=2).comment = Comment(
        "This is the loaded entity's own net invested capital, not necessarily what a specific "
        "LP/GP holds after the JV waterfall — a reasonable real starting point for a top-line IRR, "
        "but override if you want the actual equity basis for a specific investor.", "Deal Dashboard"
    )
    define("StartingEquity", "Assumptions", 2, r)
    r += 2

    r = _section_header(wsA, r, "Blended Market Leasing Terms (computed)", span=4)
    for label, key_suffix in [
        ("Blended Downtime (months)", "Downtime"),
        ("Blended Free Rent (months)", "FreeRent"),
        ("Blended TI ($/SF)", "TI"),
        ("Blended LC (% of Year-1 rent)", "LC"),
    ]:
        wsA.cell(row=r, column=1, value=label)
        fmt = PCT_FMT if key_suffix == "LC" else (MONEY_PSF_FMT if key_suffix == "TI" else "0.0")
        _formula_cell(wsA, r, 2, f"=RenewalProb*Renewal{key_suffix}+(1-RenewalProb)*New{key_suffix}", fmt)
        define(f"Blended{key_suffix}", "Assumptions", 2, r)
        r += 1

    wsA.freeze_panes = "A4"

    # =======================================================================
    # Sheet 2: Rent Roll (Current)
    # =======================================================================
    ws1 = out.create_sheet("Rent Roll (Current)")
    ws1.sheet_view.showGridLines = False
    _autosize(ws1, [10, 12, 10, 12, 12, 30, 16, 12, 14, 12])

    ws1["A1"] = f"{cfg.display()} — Current Rent Roll"
    ws1["A1"].font = TITLE_FONT
    ws1["A2"] = f"As of {as_of.strftime('%B %d, %Y') if as_of else 'n/a'} — source: Deal Dashboard rent roll, currently loaded period."
    ws1["A2"].font = NOTE_FONT

    r = 4
    header_row = r
    headers = ["Building", "Floor", "Unit", "SF", "Status", "Tenant", "Annual Rent", "Rent PSF", "Lease Expiration", "Rollover Year"]
    for i, h in enumerate(headers, start=1):
        ws1.cell(row=header_row, column=i, value=h)
    _style_header_row(ws1, header_row, len(headers))
    r += 1
    rent_roll_rows = {}
    for idx, line in enumerate(rent_roll.lines):
        ws1.cell(row=r, column=1, value=line.building).border = BORDER
        ws1.cell(row=r, column=2, value=str(line.floor)).border = BORDER
        ws1.cell(row=r, column=3, value=line.unit_code).border = BORDER
        c4 = ws1.cell(row=r, column=4, value=line.unit_area); c4.number_format = SF_FMT; c4.border = BORDER
        ws1.cell(row=r, column=5, value="VACANT" if line.is_vacant else "Leased").border = BORDER
        ws1.cell(row=r, column=6, value=line.tenant_name or "").border = BORDER
        c7 = ws1.cell(row=r, column=7, value=line.annual_rent); c7.number_format = MONEY_FMT; c7.border = BORDER
        c8 = ws1.cell(row=r, column=8, value=line.annual_rent_psf); c8.number_format = MONEY_PSF_FMT; c8.border = BORDER
        ws1.cell(row=r, column=9, value=line.lease_to.strftime("%m/%d/%Y") if line.lease_to else "").border = BORDER
        ry = rollover_year(line)
        ryc = ws1.cell(row=r, column=10, value=ry); ryc.border = BORDER
        ryc.comment = Comment(
            "Computed from the real lease expiration date (years remaining, rounded up; vacant "
            "suites = Year 1). Simplification: models each suite's FIRST rollover only.", "Deal Dashboard"
        )
        for c in range(1, 11):
            ws1.cell(row=r, column=c).font = FORMULA_FONT
        rent_roll_rows[idx] = (r, line, ry)
        r += 1
    ws1.freeze_panes = "A5"

    # =======================================================================
    # Sheet 3: Rollover & Rent Projection
    # =======================================================================
    ws3 = out.create_sheet("Rollover & Rent Projection")
    ws3.sheet_view.showGridLines = False
    _autosize(ws3, [34] + [14] * (N_YEARS + 1))

    ws3["A1"] = "Rollover & Rent Projection"
    ws3["A1"].font = TITLE_FONT
    ws3["A2"] = (
        "Pre-rollover years hold each suite's real current rent flat. At rollover, rent resets to "
        "the escalated market rate below; downtime/free rent reduce cash collected only in the "
        "rollover year itself."
    )
    ws3["A2"].font = NOTE_FONT
    ws3.merge_cells("A2:F2")
    ws3.row_dimensions[2].height = 28

    r = 4
    _year_header_row(ws3, r)
    r += 1

    ws3.cell(row=r, column=1, value="Market Rent, $/SF/Yr (escalated)").font = BOLD_FONT
    market_rent_row = r
    _formula_cell(ws3, r, 3, "=MarketRentY1", MONEY_PSF_FMT)
    for i, col in enumerate(YEAR_COLS[1:], start=1):
        prev_col = YEAR_COLS[i - 1]
        _formula_cell(ws3, r, col, f"={_col(prev_col)}{r}*(1+RentGrowth)", MONEY_PSF_FMT)
    r += 2

    ws3.cell(row=r, column=1, value="Per-Suite Contract Rent ($/Yr)").font = SECTION_FONT
    r += 1
    contract_rows = {}
    cash_rows = {}
    for idx, (rr_row, line, ry) in rent_roll_rows.items():
        label = f"{line.building}-{line.floor}-{line.unit_code}" + (" (vacant)" if line.is_vacant else f" — {line.tenant_name}")
        ws3.cell(row=r, column=1, value=f"Contract Rent: {label}")
        sf_ref = f"'Rent Roll (Current)'!D{rr_row}"
        cur_rent_ref = f"'Rent Roll (Current)'!G{rr_row}"
        for col in YEAR_COLS:
            year_num = col - YEAR_COLS[0] + 1
            cl = _col(col)
            formula = f"={cur_rent_ref}" if year_num < ry else f"={sf_ref}*{cl}${market_rent_row}"
            _formula_cell(ws3, r, col, formula, MONEY_FMT)
        contract_rows[idx] = r
        r += 1
    r += 1

    ws3.cell(row=r, column=1, value="Per-Suite Cash Rent ($/Yr, net of rollover downtime & free rent)").font = SECTION_FONT
    r += 1
    for idx, (rr_row, line, ry) in rent_roll_rows.items():
        label = f"{line.building}-{line.floor}-{line.unit_code}" + (" (vacant)" if line.is_vacant else f" — {line.tenant_name}")
        ws3.cell(row=r, column=1, value=f"Cash Rent: {label}")
        crow = contract_rows[idx]
        for col in YEAR_COLS:
            year_num = col - YEAR_COLS[0] + 1
            cl = _col(col)
            if year_num == ry:
                formula = f"={cl}{crow}*(12-BlendedDowntime-BlendedFreeRent)/12"
            else:
                formula = f"={cl}{crow}"
            _formula_cell(ws3, r, col, formula, MONEY_FMT)
        cash_rows[idx] = r
        r += 1
    r += 1

    ws3.cell(row=r, column=1, value="Total Contract Rent").font = BOLD_FONT
    total_contract_row = r
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws3, r, col, "=" + "+".join(f"{cl}{cr}" for cr in contract_rows.values()), MONEY_FMT, bold=True)
    r += 1
    ws3.cell(row=r, column=1, value="Total Cash Rent (Gross Potential Rent)").font = BOLD_FONT
    total_cash_row = r
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws3, r, col, "=" + "+".join(f"{cl}{cr}" for cr in cash_rows.values()), MONEY_FMT, bold=True)
    r += 1
    ws3.cell(row=r, column=1, value="Rent Collection Factor (Cash / Contract)")
    ws3.cell(row=r, column=1).comment = Comment(
        "Used to scale Recovery Income proportionally in rollover years.", "Deal Dashboard")
    collection_factor_row = r
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws3, r, col, f"={cl}{total_cash_row}/{cl}{total_contract_row}", PCT_FMT)
    r += 2

    ws3.cell(row=r, column=1, value="TI / LC Cost This Year ($)").font = SECTION_FONT
    ws3.cell(row=r, column=1).comment = Comment(
        "Fires only in each suite's own rollover year: SF x Blended TI, plus Blended LC% x that "
        "suite's Year-1-post-rollover contract rent (a simplified proxy for total lease value, "
        "since no Lease Term assumption is modeled). Escalates by TI/LC Growth for rollovers "
        "happening in later years.", "Deal Dashboard")
    r += 1
    ti_lc_row_start = r
    for idx, (rr_row, line, ry) in rent_roll_rows.items():
        label = f"{line.building}-{line.floor}-{line.unit_code}"
        ws3.cell(row=r, column=1, value=f"TI/LC: {label}")
        sf_ref = f"'Rent Roll (Current)'!D{rr_row}"
        crow = contract_rows[idx]
        for col in YEAR_COLS:
            year_num = col - YEAR_COLS[0] + 1
            cl = _col(col)
            if year_num == ry:
                escal = f"*(1+TILCGrowth)^{ry - 1}"
                formula = f"=({sf_ref}*BlendedTI{escal})+(BlendedLC*{cl}{crow})"
            else:
                formula = "=0"
            _formula_cell(ws3, r, col, formula, MONEY_FMT)
        r += 1
    ti_lc_row_end = r - 1
    ws3.cell(row=r, column=1, value="Total TI / LC Cost").font = BOLD_FONT
    total_ti_lc_row = r
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws3, r, col, f"=SUM({cl}{ti_lc_row_start}:{cl}{ti_lc_row_end})", MONEY_FMT, bold=True)

    ws3.freeze_panes = "B5"

    # =======================================================================
    # Sheet 4: Cash Flow Projection
    # =======================================================================
    ws4 = out.create_sheet("Cash Flow Projection")
    ws4.sheet_view.showGridLines = False
    _autosize(ws4, [40] + [14] * (N_YEARS + 1))

    ws4["A1"] = "Cash Flow Projection"
    ws4["A1"].font = TITLE_FONT
    ws4["A2"] = "Shows all 10 modeled years regardless of Hold Period — the Returns sheet applies the Hold Period gate and the sale."
    ws4["A2"].font = NOTE_FONT
    ws4.merge_cells("A2:F2")

    r = 4
    _year_header_row(ws4, r)
    r += 2

    def year0_and_years(label, year0_formula, year_formula_fn, fmt=MONEY_FMT, bold=False):
        nonlocal r
        ws4.cell(row=r, column=1, value=label)
        if bold:
            ws4.cell(row=r, column=1).font = BOLD_FONT
        if year0_formula is not None:
            _formula_cell(ws4, r, 2, year0_formula, fmt, bold=bold)
        for i, col in enumerate(YEAR_COLS, start=1):
            _formula_cell(ws4, r, col, year_formula_fn(i, col), fmt, bold=bold)
        this_row = r
        r += 1
        return this_row

    gpr_row = year0_and_years(
        "Gross Potential Rent (Total Cash Rent)", None,
        lambda i, col: f"='Rollover & Rent Projection'!{_col(col)}{total_cash_row}",
    )
    vac_credit_row = year0_and_years(
        "Less: Vacancy & Credit Loss", None,
        lambda i, col: f"=-{_col(col)}{gpr_row}*VacancyCredit",
    )
    egi_row = year0_and_years(
        "Effective Rental Income", f"=B{gpr_row}",
        lambda i, col: f"={_col(col)}{gpr_row}+{_col(col)}{vac_credit_row}", bold=True,
    )

    revenue_actual = annual["revenue"].actual_value if annual.get("revenue") else None
    expenses_actual = annual["expenses"].actual_value if annual.get("expenses") else None
    nonop_actual = annual["non_operating_expenses"].actual_value if annual.get("non_operating_expenses") else None
    capex_budget = annual["capital_expenditures"].budget_value if annual.get("capital_expenditures") else None

    # Opex/Non-op computed before Recovery Income so Recovery can reference the Opex row. Each
    # grows off its own previous column (col-1 == column B == Year 0 for the first projection
    # year). The Year-0 baseline and every subsequent value here is already negative (an
    # expense), so growing it is `prev * (1+growth)`, not `-prev * (1+growth)` (that second form
    # would flip the sign every year).
    opex_row = year0_and_years(
        "Operating Expenses", f"=-{expenses_actual or 0}/{months_ytd}*12",
        lambda i, col: f"={_col(col-1)}{r}*(1+OpexGrowth)",
    )
    nonop_row = year0_and_years(
        "Non-operating Expenses", f"=-{nonop_actual or 0}/{months_ytd}*12",
        lambda i, col: f"={_col(col-1)}{r}*(1+OpexGrowth)",
    )
    recovery_row = year0_and_years(
        "Recovery Income", f"=RecoveryPct*-B{opex_row}",
        lambda i, col: (
            f"=RecoveryPct*-{_col(col)}{opex_row}*"
            f"'Rollover & Rent Projection'!{_col(col)}{collection_factor_row}"
        ),
    )

    total_rev_row = year0_and_years(
        "Total Revenue", f"=B{egi_row}+B{recovery_row}",
        lambda i, col: f"={_col(col)}{egi_row}+{_col(col)}{recovery_row}", bold=True,
    )
    noi_row = year0_and_years(
        "NOI", f"=B{total_rev_row}+B{opex_row}+B{nonop_row}",
        lambda i, col: f"={_col(col)}{total_rev_row}+{_col(col)}{opex_row}+{_col(col)}{nonop_row}", bold=True,
    )
    ti_lc_row = year0_and_years(
        "Less: TI / LC", None,
        lambda i, col: f"=-'Rollover & Rent Projection'!{_col(col)}{total_ti_lc_row}",
    )
    capex_row = year0_and_years(
        "Less: Capital Expenditures", f"=-{capex_budget or 0}",
        lambda i, col: f"={_col(col-1)}{r}*(1+CapitalGrowth)",
    )
    ws4.cell(row=capex_row, column=2).comment = Comment(
        "Year 0 uses the real full-year Capital Expenditures BUDGET figure (not YTD-annualized — "
        "capex isn't evenly spread through the year). Escalated by Capital Growth thereafter as a "
        "placeholder; replace with a real multi-year CapEx plan if/when one is dropped into the app.",
        "Deal Dashboard",
    )
    cfbds_row = year0_and_years(
        "Cash Flow Before Debt Service", f"=B{noi_row}+B{ti_lc_row}+B{capex_row}",
        lambda i, col: f"={_col(col)}{noi_row}+{_col(col)}{ti_lc_row}+{_col(col)}{capex_row}", bold=True,
    )

    r += 1
    ws4.cell(row=r, column=1, value="Debt Schedule").font = SECTION_FONT
    r += 1
    ws4.cell(row=r, column=1, value=(
        "Pre-refinance debt is modeled interest-only (real debt service ties closely to pure "
        "interest on the balance, see Assumptions). Post-refinance uses the new loan's own "
        "amortization via PMT/PPMT."
    )).font = NOTE_FONT
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    ws4.row_dimensions[r].height = 28
    r += 1

    beginning_row = year0_and_years("Beginning Balance", "=ExistingDebtBalance", lambda i, col: None)
    rate_row = year0_and_years(
        "Rate This Year", None,
        lambda i, col: f"=IF(AND(RefiYear>0,{i}>=RefiYear),RefiRate,ExistingBlendedRate)", fmt=PCT1_FMT,
    )
    years_since_refi_row = year0_and_years(
        "Years Since Refinance (post-refi only)", None,
        lambda i, col: f"=IF(AND(RefiYear>0,{i}>=RefiYear),{i}-RefiYear+1,0)", fmt="0",
    )
    io_flag_row = year0_and_years(
        "Interest-Only This Year?", None,
        lambda i, col: f"=IF({_col(col)}{years_since_refi_row}>0,IF({_col(col)}{years_since_refi_row}<=RefiIOYears,1,0),0)", fmt="0",
    )
    interest_row = year0_and_years(
        "Interest Payment", f"=B{beginning_row}*ExistingBlendedRate",
        lambda i, col: f"={_col(col)}{beginning_row}*{_col(col)}{rate_row}",
    )
    principal_row = year0_and_years(
        "Principal Payment", "=0",
        lambda i, col: (
            f"=IF({_col(col)}{io_flag_row}=1,0,"
            f"IF(AND(RefiYear>0,{i}>=RefiYear),"
            f"PPMT(RefiRate,{_col(col)}{years_since_refi_row}-RefiIOYears,RefiAmortYears,-RefiAmount),0))"
        ),
    )
    ending_row = year0_and_years(
        "Ending Balance", f"=B{beginning_row}",
        lambda i, col: f"={_col(col)}{beginning_row}-{_col(col)}{principal_row}",
    )
    for i, col in enumerate(YEAR_COLS, start=1):
        cl, prev_cl = _col(col), _col(col - 1)
        formula = (
            f"=IF(AND(RefiYear>0,{i}>=RefiYear),IF({i}=RefiYear,RefiAmount,{prev_cl}{ending_row}),"
            f"IF({i}=1,ExistingDebtBalance,{prev_cl}{ending_row}))"
        )
        _formula_cell(ws4, beginning_row, col, formula, MONEY_FMT)
    for row_ in (beginning_row, interest_row, principal_row, ending_row):
        for col in [2] + YEAR_COLS:
            ws4.cell(row=row_, column=col).number_format = MONEY_FMT

    total_ds_row = year0_and_years(
        "Total Debt Service", f"=B{interest_row}+B{principal_row}",
        lambda i, col: f"={_col(col)}{interest_row}+{_col(col)}{principal_row}", bold=True,
    )
    levered_cf_row = year0_and_years(
        "Levered Cash Flow (before sale)", f"=B{cfbds_row}-B{total_ds_row}",
        lambda i, col: f"={_col(col)}{cfbds_row}-{_col(col)}{total_ds_row}", bold=True,
    )

    ws4.freeze_panes = "B5"

    # =======================================================================
    # Sheet 5: Returns
    # =======================================================================
    ws5 = out.create_sheet("Returns")
    ws5.sheet_view.showGridLines = False
    _autosize(ws5, [40] + [14] * (N_YEARS + 1))

    ws5["A1"] = "Returns"
    ws5["A1"].font = TITLE_FONT
    ws5["A2"] = (
        "Applies the Hold Period gate to the Cash Flow Projection sheet and adds net sale proceeds "
        "in the exit year. Exit uses the exit year's own (trailing) NOI x Cap Rate, not forward NOI."
    )
    ws5["A2"].font = NOTE_FONT
    ws5.merge_cells("A2:F2")

    r = 4
    _year_header_row(ws5, r)
    r += 2

    ws5.cell(row=r, column=1, value="Is Exit Year?")
    for i, col in enumerate(YEAR_COLS, start=1):
        _formula_cell(ws5, r, col, f"=IF({i}=HoldPeriod,1,0)", "0")
    exit_flag_row = r
    r += 1

    ws5.cell(row=r, column=1, value="Exit Year NOI (trailing)")
    for i, col in enumerate(YEAR_COLS, start=1):
        cl = _col(col)
        _formula_cell(ws5, r, col, f"=IF({cl}{exit_flag_row}=1,'Cash Flow Projection'!{cl}{noi_row},0)", MONEY_FMT)
    exit_noi_row = r
    r += 1

    ws5.cell(row=r, column=1, value="Gross Sale Value")
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws5, r, col, f"=IF({cl}{exit_flag_row}=1,{cl}{exit_noi_row}/CapRate,0)", MONEY_FMT)
    gross_sale_row = r
    r += 1

    ws5.cell(row=r, column=1, value="Less: Cost of Sale")
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws5, r, col, f"=-{cl}{gross_sale_row}*CostOfSale", MONEY_FMT)
    cost_sale_row = r
    r += 1

    ws5.cell(row=r, column=1, value="Less: Remaining Debt Balance at Exit")
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws5, r, col, f"=IF({cl}{exit_flag_row}=1,-'Cash Flow Projection'!{cl}{ending_row},0)", MONEY_FMT)
    payoff_row = r
    r += 1

    ws5.cell(row=r, column=1, value="Net Sale Proceeds").font = BOLD_FONT
    for col in YEAR_COLS:
        cl = _col(col)
        _formula_cell(ws5, r, col, f"={cl}{gross_sale_row}+{cl}{cost_sale_row}+{cl}{payoff_row}", MONEY_FMT, bold=True)
    net_sale_row = r
    r += 2

    ws5.cell(row=r, column=1, value="Cash Flow to Equity").font = BOLD_FONT
    _formula_cell(ws5, r, 2, "=-StartingEquity", MONEY_FMT, bold=True)
    for i, col in enumerate(YEAR_COLS, start=1):
        cl = _col(col)
        formula = f"=IF({i}<=HoldPeriod,'Cash Flow Projection'!{cl}{levered_cf_row}+{cl}{net_sale_row},0)"
        _formula_cell(ws5, r, col, formula, MONEY_FMT, bold=True)
    cf_to_equity_row = r
    r += 2

    ws5.cell(row=r, column=1, value="IRR")
    _formula_cell(ws5, r, 2, f"=IRR(B{cf_to_equity_row}:{_col(YEAR_COLS[-1])}{cf_to_equity_row})", PCT_FMT, bold=True)
    r += 1
    ws5.cell(row=r, column=1, value="Equity Multiple")
    _formula_cell(
        ws5, r, 2,
        f'=SUMIF(C{cf_to_equity_row}:{_col(YEAR_COLS[-1])}{cf_to_equity_row},">0")/-B{cf_to_equity_row}',
        MULT_FMT, bold=True,
    )

    ws5.freeze_panes = "B5"

    for name, ref in names.items():
        out.defined_names[name] = DefinedName(name, attr_text=ref)

    return out
