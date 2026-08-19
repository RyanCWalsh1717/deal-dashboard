"""Property detail view: Summary / Cash / Equity & Capital / Debt & Loans /
Distribution Waterfall (per-tier tabs, each with Current + Stabilized) /
Budget vs. Actuals (Summary + Detailed) / Sources & Uses (placeholder) /
Leasing & Investment Outlook (placeholder).
"""

from __future__ import annotations

import io
import re
from typing import List, Optional

import pandas as pd
import streamlit as st

from pipeline.models import (
    CashAccountBalance,
    DistributionWorkbookResult,
    EntityTrialBalance,
    LoanStatement,
    RentRollResult,
    WaterfallTier,
)
from pipeline import holdsell_model
from pipeline.parsers.abstract_loader import load_jv_abstract, load_loan_abstract
from pipeline.property_config import PropertyConfig
from views.branding import render_hero, render_kpi_tiles


def _fmt_money(v):
    if v is None:
        return "—"
    return f"-${abs(v):,.0f}" if v < 0 else f"${v:,.0f}"


def _money_col(label: Optional[str] = None):
    """A numeric $-formatted column (no decimals) for st.dataframe's
    column_config — keeps the underlying value numeric so it right-aligns
    and sorts correctly, instead of a pre-formatted string.

    Uses the "dollar" preset (Intl.NumberFormat currency style) rather than
    a "$%,.0f" printf string — printf just prepends the literal "$" before
    whatever sprintf produces, so negative values render as "$-1,234"
    instead of "-$1,234". The "dollar" preset formats the sign correctly.
    `step=1` is what tells Streamlit's formatter to use 0 decimal places
    (it derives display precision from the column's step, not the dtype).
    """
    return st.column_config.NumberColumn(label, format="dollar", step=1)


def _pct_col(label: Optional[str] = None):
    """A numeric %-formatted column (1 decimal). Expects the underlying
    value already scaled to 0-100 (not a 0-1 fraction) — see _pct100()."""
    return st.column_config.NumberColumn(label, format="%.1f%%")


def _pct100(v: Optional[float]) -> Optional[float]:
    return None if v is None else v * 100


def _fmt_pct(v):
    return "—" if v is None else f"{v * 100:.1f}%"


def _total_cash(result: Optional[DistributionWorkbookResult]) -> Optional[float]:
    if not result or not result.equity:
        return None
    values = [pos.total_cash for pos in result.equity.values() if pos.total_cash is not None]
    return sum(values) if values else None


def _cash_boxes(
    result: Optional[DistributionWorkbookResult], cash_accounts: List[CashAccountBalance]
) -> List[tuple]:
    """Every individual cash-bearing account the app currently has data for,
    across every entity level: JV/venture operating cash from the equity
    tabs, plus trial-balance/loan-statement-derived development cash and
    escrows/reserves (which live at the property-entity level and are a
    materially different, non-overlapping pool of money). Deduped by
    (entity_code, label) rather than label alone — once multiple entities'
    trial balances are loaded, two different entities can legitimately both
    have a "Cash - Operating" line, and deduping on label only would wrongly
    collapse them into one."""
    boxes = [("Operating Cash (all entities)", _total_cash(result), "equity_tabs")]
    seen = {("", boxes[0][0])}
    for acct in cash_accounts:
        key = (acct.entity_code, acct.label)
        if key in seen:
            continue
        seen.add(key)
        boxes.append((acct.label, acct.balance, acct.source))
    return boxes


def _total_cash_all_sources(
    result: Optional[DistributionWorkbookResult], cash_accounts: List[CashAccountBalance]
) -> Optional[float]:
    values = [v for _, v, _ in _cash_boxes(result, cash_accounts) if v is not None]
    return sum(values) if values else None


def _current_debt(
    result: Optional[DistributionWorkbookResult], loan_statements: List[LoanStatement]
) -> tuple:
    """Prefers actual lender-statement balances over the distribution
    workbook's pre-paydown forecast, since the two can diverge materially
    (confirmed: forecast assumes a paydown that hadn't posted yet) — returns
    (total_outstanding, source_label) so callers can show which one they got."""
    if loan_statements:
        total = sum(s.principal_balance or 0 for s in loan_statements)
        as_of_dates = {s.as_of for s in loan_statements if s.as_of}
        sub = f"actual, as of {max(as_of_dates)}" if as_of_dates else "actual, per lender"
        return total, sub
    if result and result.debt and result.debt.tranches:
        return result.debt.total_outstanding, "forecast (distribution workbook)"
    return None, None


SECTIONS = [
    "Summary",
    "Cash",
    "Equity & Capital",
    "Balance Sheet",
    "Rent Roll",
    "Debt & Loans",
    "Reconciliation",
    "Distribution Waterfall",
    "Budget vs. Actuals",
    "Annual Budget",
    "Hold/Sell Assumptions",
    "Sources & Uses",
    "Leasing & Investment Outlook",
]


def _goto(section: str, budget_subtab: Optional[str] = None) -> None:
    """Programmatic navigation. Can't write directly to `st.session_state.detail_section`
    here — by the time a Summary-tab button is clicked, the segmented_control widget
    bound to that key has already been instantiated earlier in this same run, and
    Streamlit forbids mutating a widget's key after that. Instead, stage the target
    in separate variables that get applied at the top of the next run, BEFORE the
    widgets are (re)created."""
    st.session_state.pending_section = section
    if budget_subtab:
        st.session_state.pending_budget_subtab = budget_subtab
    st.rerun()


def render_property_detail(
    cfg: PropertyConfig,
    result: Optional[DistributionWorkbookResult],
    cash_accounts: Optional[List[CashAccountBalance]] = None,
    rent_roll: Optional[RentRollResult] = None,
    loan_statements: Optional[List[LoanStatement]] = None,
    entity_trial_balances: Optional[List[EntityTrialBalance]] = None,
    data_dir: str = "data",
    opex_categories: Optional[dict] = None,
) -> None:
    cash_accounts = cash_accounts or []
    loan_statements = loan_statements or []
    entity_trial_balances = entity_trial_balances or []
    opex_categories = opex_categories or {}

    badges = [b for b in [cfg.market, cfg.property_type] if b]
    render_hero(cfg.display(), cfg.property_address, badges, photo_code=cfg.property_code)
    st.write("")

    # Apply any pending nav request (from a Summary jump button) before either
    # segmented_control widget below is instantiated this run.
    if "pending_section" in st.session_state:
        st.session_state.detail_section = st.session_state.pop("pending_section")
    if "pending_budget_subtab" in st.session_state:
        st.session_state.budget_subtab = st.session_state.pop("pending_budget_subtab")

    if "detail_section" not in st.session_state or st.session_state.detail_section not in SECTIONS:
        st.session_state.detail_section = SECTIONS[0]
    section = st.segmented_control("Section", SECTIONS, key="detail_section", label_visibility="collapsed")
    if section is None:  # user clicked the active pill again and deselected it
        st.session_state.detail_section = SECTIONS[0]
        section = SECTIONS[0]

    st.divider()

    if section == "Summary":
        _render_summary(cfg, result, rent_roll, cash_accounts, loan_statements)
    elif section == "Cash":
        _render_cash(result, cash_accounts)
    elif section == "Equity & Capital":
        _render_equity(result)
    elif section == "Balance Sheet":
        _render_balance_sheet(result)
    elif section == "Rent Roll":
        _render_rent_roll(rent_roll)
    elif section == "Debt & Loans":
        _render_debt(cfg, result, loan_statements)
    elif section == "Reconciliation":
        _render_reconciliation(cfg, entity_trial_balances, loan_statements, result)
    elif section == "Distribution Waterfall":
        _render_waterfall(cfg, result)
    elif section == "Budget vs. Actuals":
        _render_budget(result)
    elif section == "Annual Budget":
        _render_annual_budget(result)
    elif section == "Hold/Sell Assumptions":
        _render_holdsell_assumptions(cfg, result, rent_roll, loan_statements, entity_trial_balances, data_dir, opex_categories)
    elif section == "Sources & Uses":
        st.info(
            "Sources & Uses will populate once the leasing/investment outlook model is finalized — "
            "it lives entirely in that model, not the distribution workbook."
        )
    elif section == "Leasing & Investment Outlook":
        st.info(
            "Leasing & Investment Outlook integration is not yet available — the source model is "
            "still in development. This section will populate once that workbook is finalized."
        )


def _render_summary(
    cfg: PropertyConfig,
    result: Optional[DistributionWorkbookResult],
    rent_roll: Optional[RentRollResult],
    cash_accounts: Optional[List[CashAccountBalance]] = None,
    loan_statements: Optional[List[LoanStatement]] = None,
) -> None:
    cash_accounts = cash_accounts or []
    loan_statements = loan_statements or []

    total_cash = _total_cash_all_sources(result, cash_accounts)
    debt_total, debt_sub = _current_debt(result, loan_statements)
    noi_today = None
    noi_period = None
    if result and result.budget_summary:
        noi_line = next((l for l in result.budget_summary.lines if l.account_code == "noi"), None)
        if noi_line:
            noi_today = noi_line.actual_value
        noi_period = result.budget_summary.period

    noi_sub = f"as of {noi_period.strftime('%B %Y')}" if noi_period else None
    render_kpi_tiles(
        [
            ("Current Cash", _fmt_money(total_cash), "all accounts, all entities"),
            ("NOI (Last Actual Month)", _fmt_money(noi_today), noi_sub),
            ("Total Debt Outstanding", _fmt_money(debt_total), debt_sub),
            ("Stabilized NOI", "—", "pending leasing model"),
        ]
    )
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("View Cash →", key="goto_cash"):
            _goto("Cash")
    with col2:
        if st.button("View Detailed Breakdown →", key="goto_noi_detail"):
            _goto("Budget vs. Actuals", budget_subtab="Detailed")
    with col3:
        if st.button("View Debt & Loans →", key="goto_debt_top"):
            _goto("Debt & Loans")
    if noi_period:
        st.caption(f"NOI as of {noi_period.strftime('%B %Y')} — the last month the Cash Flow tab labels as Actuals.")
    st.caption("Stabilized NOI will populate once the leasing/investment outlook model is finalized.")

    st.divider()
    st.markdown("#### Current Cash — by Account")
    cash_rows = [
        {"Account": label, "Balance": value}
        for label, value, _source in _cash_boxes(result, cash_accounts)
    ]
    if cash_rows:
        st.dataframe(
            pd.DataFrame(cash_rows),
            width="stretch",
            hide_index=True,
            column_config={"Balance": _money_col()},
        )
        st.caption(
            "Includes JV/venture-level operating cash (from the distribution workbook's equity tabs) "
            "and property-level development cash + escrows/reserves (from the trial balance or loan "
            "statements) — these are different, non-overlapping entities' cash, not duplicates."
        )
    else:
        st.info("No cash account data available yet.")

    st.divider()
    st.markdown("#### Loan Terms")
    if loan_statements:
        df = pd.DataFrame(
            [
                {"Tranche": s.tranche_name, "Balance": s.principal_balance, "Rate": _pct100(s.interest_rate)}
                for s in loan_statements
            ]
        )
        st.dataframe(
            df,
            width="stretch",
            hide_index=True,
            column_config={"Balance": _money_col(), "Rate": _pct_col()},
        )
        st.caption(f"Actual balances and all-in rates, per lender statements ({debt_sub}).")
    elif result and result.debt and result.debt.tranches:
        df = pd.DataFrame(
            [
                {"Tranche": t.tranche_name, "Balance": t.outstanding_balance, "Rate": _pct100(t.interest_rate)}
                for t in result.debt.tranches
            ]
        )
        st.dataframe(
            df,
            width="stretch",
            hide_index=True,
            column_config={"Balance": _money_col(), "Rate": _pct_col()},
        )
        st.caption(
            "Forecast from the distribution workbook — rate shown is the spread over SOFR, not the "
            "all-in rate, and the balance assumes a paydown that may not have posted yet. No loan "
            "statements uploaded yet for actuals."
        )
    else:
        st.info("No loan data available.")

    st.divider()
    st.markdown("#### Sources & Uses")
    st.info("Will populate once the leasing/investment outlook model is finalized.")

    st.divider()
    st.markdown("#### Rent Roll")
    if not rent_roll or not rent_roll.lines:
        st.info("Not yet provided — upload a Tenancy Schedule export in the sidebar (Property Detail view).")
    else:
        if rent_roll.as_of:
            st.caption(f"As of {rent_roll.as_of.strftime('%B %d, %Y')}")
        render_kpi_tiles(
            [
                ("Occupancy", _fmt_pct(rent_roll.occupancy_pct), None),
                ("Leased SF", f"{rent_roll.total_leased_sf:,.0f}", None),
                ("Vacant SF", f"{rent_roll.total_vacant_sf:,.0f}", None),
            ]
        )
        if st.button("View Rent Roll →", key="goto_rent_roll"):
            _goto("Rent Roll")


def _render_cash(result: Optional[DistributionWorkbookResult], cash_accounts: List[CashAccountBalance]) -> None:
    boxes = _cash_boxes(result, cash_accounts)
    boxes.insert(1, ("DACA", None, "placeholder"))  # display-only — no source yet, excluded from any total

    render_kpi_tiles([(label, _fmt_money(value), None) for label, value, _source in boxes])

    if not cash_accounts:
        st.caption(
            "Escrow/reserve boxes will appear here once you upload a trial balance or loan statement "
            "in the sidebar (Property Detail view). DACA has no source yet — placeholder for now."
        )

    if not result or not result.cash_flow or not result.cash_flow.lines:
        return

    st.divider()
    cf = result.cash_flow
    latest = cf.latest_period()
    latest_total = sum(line.monthly_values.get(latest, 0.0) for line in cf.lines) if latest else None
    render_kpi_tiles([("Latest Month Net Activity", _fmt_money(latest_total), None)])

    periods = cf.period_columns[-12:]
    totals = {p: sum(line.monthly_values.get(p, 0.0) for line in cf.lines) for p in periods}
    df = pd.DataFrame({"Net Activity": totals})
    st.line_chart(df)
    st.caption("Monthly net revenue/expense activity from the Cash Flow tab (trailing 12 months).")


def _render_equity(result: Optional[DistributionWorkbookResult]) -> None:
    if not result or not result.equity:
        st.info("No equity data available.")
        return

    for key, pos in result.equity.items():
        st.markdown(f"### {pos.entity_name or key}")
        render_kpi_tiles(
            [
                ("Cash Balance", _fmt_money(pos.total_cash), None),
                ("Total Contributions", _fmt_money(pos.total_contributions), None),
                ("Total Distributions", _fmt_money(pos.total_distributions), None),
            ]
        )

        partners = sorted(set(pos.contributions_by_partner) | set(pos.distributions_by_partner))
        if partners:
            df = pd.DataFrame(
                [
                    {
                        "Partner": p,
                        "Contributions": pos.contributions_by_partner.get(p),
                        "Distributions": pos.distributions_by_partner.get(p),
                    }
                    for p in partners
                ]
            )
            st.dataframe(
                df,
                width="stretch",
                hide_index=True,
                column_config={"Contributions": _money_col(), "Distributions": _money_col()},
            )
        st.divider()

    st.markdown("#### Future Distributions")
    st.info("Will populate once the leasing/investment outlook model is finalized.")


def _render_balance_sheet(result: Optional[DistributionWorkbookResult]) -> None:
    if not result or not result.equity:
        st.info("No balance sheet data available.")
        return

    for key, pos in result.equity.items():
        if not pos.balance_sheet_lines:
            continue

        st.markdown(f"### {pos.entity_name or key}")
        if pos.as_of_period:
            st.caption(f"Period: {pos.as_of_period}")

        # A line with no value that shares an account code with a valued line
        # further down is just the pre-total header for that same line (e.g.
        # "Cash - Operating - BofA" immediately before "Total Cash - Operating
        # - BofA", both code 111110) — skip the redundant blank one. Section
        # headers (ASSETS, CASH & CASH EQUIVALENTS, ...) have no such
        # same-code total elsewhere, so they're untouched by this filter.
        valued_codes = {l.account_code for l in pos.balance_sheet_lines if l.value is not None}

        rows = []
        for line in pos.balance_sheet_lines:
            if line.value is None and line.account_code in valued_codes:
                continue
            indent_str = "  " * line.indent
            rows.append(
                {
                    "Account": line.account_code,
                    "Line Item": f"{indent_str}{line.label}",
                    "Balance": line.value,
                }
            )

        st.dataframe(
            pd.DataFrame(rows),
            width="stretch",
            hide_index=True,
            column_config={"Balance": _money_col()},
        )
        st.divider()


def _render_rent_roll(rent_roll: Optional[RentRollResult]) -> None:
    if not rent_roll or not rent_roll.lines:
        st.info("Not yet provided — upload a Tenancy Schedule export in the sidebar (Property Detail view).")
        return

    as_of = rent_roll.as_of
    if as_of:
        st.caption(f"As of {as_of.strftime('%B %d, %Y')}")

    walt = rent_roll.weighted_average_lease_term_years(as_of) if as_of else None
    avg_rent = rent_roll.avg_annual_rent_psf
    avg_cam = rent_roll.avg_annual_cam_psf
    render_kpi_tiles(
        [
            ("% Leased", _fmt_pct(rent_roll.occupancy_pct), None),
            ("WALT", f"{walt:.1f} yrs" if walt is not None else "—", "SF-weighted"),
            ("Avg Rent", f"${avg_rent:,.2f}/SF" if avg_rent is not None else "—", None),
            ("Avg CAM", f"${avg_cam:,.2f}/SF" if avg_cam is not None else "—", None),
        ]
    )
    render_kpi_tiles(
        [
            ("Leased SF", f"{rent_roll.total_leased_sf:,.0f}", None),
            ("Vacant SF", f"{rent_roll.total_vacant_sf:,.0f}", None),
            ("Total Annual Rent", _fmt_money(rent_roll.total_annual_rent), None),
            ("Total Annual CAM", _fmt_money(rent_roll.total_annual_cam), None),
        ]
    )

    st.divider()
    rows = []
    for line in rent_roll.lines:
        next_step = line.next_rent_step(as_of) if as_of else None
        rows.append(
            {
                "Tenant": line.tenant_name if not line.is_vacant else "VACANT",
                "Unit": line.unit_code,
                "SF": line.lease_area or line.unit_area,
                "Lease Expiration": line.lease_to,
                "Current Rent": line.annual_rent,
                "Current CAM": line.current_annual_cam,
                "Total Obligation": line.current_total_obligation,
                "LOC/SD": line.loc_amount,
                "Next Step Date": next_step.effective_date if next_step else None,
                "Next Step Rent": next_step.annual_rent if next_step else None,
            }
        )
    st.dataframe(
        pd.DataFrame(rows),
        width="stretch",
        hide_index=True,
        column_config={
            "SF": st.column_config.NumberColumn(format="%,d"),
            "Current Rent": _money_col(),
            "Current CAM": _money_col(),
            "Total Obligation": _money_col(),
            "LOC/SD": _money_col(),
            "Next Step Rent": _money_col(),
        },
    )
    st.caption(
        '"Current CAM" sums every non-Rent charge line on the lease — see the breakdown below per '
        'tenant for the component detail (RE Tax vs. Operating recovery, where the export provides '
        "it). \"LOC/SD\" is the letter-of-credit / bank-guarantee amount on file for that lease."
    )

    st.divider()
    st.markdown("#### Escalation Schedules & Charge Detail")
    for line in rent_roll.lines:
        if line.is_vacant:
            continue
        with st.expander(f"{line.tenant_name} — Unit {line.unit_code}"):
            if line.charges:
                st.markdown("**Current charges**")
                for charge in line.charges:
                    if charge.description:
                        label = charge.description
                    elif charge.charge_type.strip().upper() == "RENT":
                        label = charge.charge_type
                    else:
                        label = f"{charge.charge_type} (component not specified in source file)"
                    st.write(f"{label}: {_fmt_money(charge.annual_amount)}")
            if line.rent_steps:
                st.markdown("**Rent escalation schedule**")
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "Effective Date": s.effective_date,
                                "Annual Rent": s.annual_rent,
                                "Rent/SF": s.annual_rent_psf,
                            }
                            for s in line.rent_steps
                        ]
                    ),
                    width="stretch",
                    hide_index=True,
                    column_config={
                        "Annual Rent": _money_col(),
                        "Rent/SF": st.column_config.NumberColumn(format="$%,.2f"),
                    },
                )
            if line.loc_amount:
                st.caption(f"LOC / Security Deposit: {_fmt_money(line.loc_amount)}")


def _normalize_tranche(name: str) -> str:
    return re.sub(r"\d+$", "", name.lower().replace(" ", ""))


def _render_debt(
    cfg: PropertyConfig,
    result: Optional[DistributionWorkbookResult],
    loan_statements: Optional[List[LoanStatement]] = None,
) -> None:
    loan_statements = loan_statements or []

    if not result or not result.debt or not result.debt.tranches:
        st.info("No debt data available.")
    else:
        debt = result.debt
        if debt.as_of:
            st.caption(f"Forecast as of {debt.as_of} (distribution workbook)")

        df = pd.DataFrame(
            [
                {
                    "Tranche": t.tranche_name,
                    "Outstanding Balance": t.outstanding_balance,
                    "Spread over SOFR": _pct100(t.interest_rate),
                }
                for t in debt.tranches
            ]
        )
        st.dataframe(
            df,
            width="stretch",
            hide_index=True,
            column_config={"Outstanding Balance": _money_col(), "Spread over SOFR": _pct_col()},
        )
        render_kpi_tiles([("Total Outstanding (Forecast)", _fmt_money(debt.total_outstanding), None)])
        st.caption(
            "This rate is the credit spread over SOFR from the distribution workbook, not the "
            "all-in borrowing cost — see Actual (Loan Statements) below for the real rate."
        )

    if loan_statements:
        st.divider()
        st.markdown("#### Actual (Loan Statements)")
        as_of_dates = {s.as_of for s in loan_statements if s.as_of}
        if as_of_dates:
            st.caption(f"As of {max(as_of_dates)} (per lender)")

        forecast_by_tranche = {
            _normalize_tranche(t.tranche_name): t
            for t in (result.debt.tranches if result and result.debt else [])
        }
        rows = []
        for stmt in loan_statements:
            forecast = forecast_by_tranche.get(_normalize_tranche(stmt.tranche_name))
            balance_delta = (
                stmt.principal_balance - forecast.outstanding_balance
                if stmt.principal_balance is not None and forecast
                else None
            )
            rows.append(
                {
                    "Tranche": stmt.tranche_name,
                    "Actual Balance": stmt.principal_balance,
                    "Actual Rate (all-in)": _pct100(stmt.interest_rate),
                    "vs. Forecast Balance": balance_delta,
                }
            )
        st.dataframe(
            pd.DataFrame(rows),
            width="stretch",
            hide_index=True,
            column_config={
                "Actual Balance": _money_col(),
                "Actual Rate (all-in)": _pct_col(),
                "vs. Forecast Balance": _money_col(),
            },
        )
        total_actual = sum(s.principal_balance or 0 for s in loan_statements)
        render_kpi_tiles([("Total Outstanding (Actual)", _fmt_money(total_actual), None)])

    if cfg.loans:
        st.markdown("#### Loan Abstracts")
        for loan_ref in cfg.loans:
            abstract = load_loan_abstract(cfg, loan_ref)
            with st.expander(loan_ref.tranche_name):
                if abstract:
                    st.json(abstract)
                else:
                    st.caption("No abstract on file yet — run `tools/extract_loan_abstract.py`.")


def _match_ownership_tier(cfg: PropertyConfig, entity_name: str):
    entity_lower = (entity_name or "").lower()
    if not entity_lower:
        return None
    for tier_cfg in cfg.ownership_tiers:
        tier_entity_lower = tier_cfg.distributing_entity.lower()
        if tier_entity_lower and (tier_entity_lower in entity_lower or entity_lower in tier_entity_lower):
            return tier_cfg
    return None


def _render_reconciliation(
    cfg: PropertyConfig,
    entity_trial_balances: List[EntityTrialBalance],
    loan_statements: List[LoanStatement],
    result: Optional[DistributionWorkbookResult],
) -> None:
    st.caption(
        "Cross-checks each uploaded trial balance's own figures against the other source files that "
        "should describe the same real-world balance — nothing here is ever combined across entities, "
        "since different levels of the JV structure hold genuinely different pools of cash/equity."
    )

    if not entity_trial_balances:
        st.info(
            'No trial balances uploaded yet — drop one in via "Update Source Files" in the sidebar to '
            "see per-entity payables, receivables, and equity, plus cross-checks against the loan "
            "statements and distribution waterfall."
        )
        return

    loan_total = sum(s.principal_balance or 0 for s in loan_statements) if loan_statements else None

    for entity in entity_trial_balances:
        st.markdown(f"### {entity.entity_name or entity.entity_code}")
        if entity.as_of:
            st.caption(f"As of {entity.as_of.strftime('%B %Y')}")

        render_kpi_tiles(
            [
                ("Accounts Receivable (Rent Outstanding)", _fmt_money(entity.accounts_receivable), None),
                ("Accounts Payable (Outstanding)", _fmt_money(entity.accounts_payable), None),
                ("Contributions", _fmt_money(entity.contributions), None),
                ("Distributions", _fmt_money(entity.distributions), None),
            ]
        )
        render_kpi_tiles(
            [
                ("Retained Earnings", _fmt_money(entity.retained_earnings), None),
                ("Mortgage Payable (per Trial Balance)", _fmt_money(entity.mortgage_payable), None),
            ]
        )

        st.markdown("**Mortgage Payable cross-check (vs. Loan Statements)**")
        if entity.mortgage_payable is not None and loan_total is not None:
            st.dataframe(
                pd.DataFrame(
                    [
                        {"Source": "Trial Balance", "Balance": entity.mortgage_payable},
                        {"Source": "Loan Statements (Actual)", "Balance": loan_total},
                        {"Source": "Variance", "Balance": entity.mortgage_payable - loan_total},
                    ]
                ),
                width="stretch",
                hide_index=True,
                column_config={"Balance": _money_col()},
            )
        elif entity.mortgage_payable is not None:
            st.caption("Mortgage Payable is on file, but no loan statements are loaded yet to cross-check against.")
        else:
            st.caption("No Mortgage Payable account found in this trial balance.")

        st.markdown("**Cash / Escrow cross-check (vs. Loan Statements)**")
        escrow_rows = []
        for acct in entity.cash_accounts:
            label_lower = acct.label.lower()
            for stmt in loan_statements:
                stmt_value = None
                if "tax escrow" in label_lower:
                    stmt_value = stmt.tax_escrow_balance
                elif "insurance escrow" in label_lower:
                    stmt_value = stmt.insurance_escrow_balance
                elif "reserve" in label_lower:
                    stmt_value = stmt.reserve_balance
                if stmt_value is not None:
                    escrow_rows.append(
                        {
                            "Account": acct.label,
                            "Trial Balance": acct.balance,
                            "Loan Statement": stmt_value,
                            "Variance": (acct.balance or 0.0) - stmt_value,
                        }
                    )
                    break
        if not loan_statements:
            st.caption("No loan statements loaded yet to cross-check cash/escrow balances against.")
        elif escrow_rows:
            st.dataframe(
                pd.DataFrame(escrow_rows),
                width="stretch",
                hide_index=True,
                column_config={
                    "Trial Balance": _money_col(),
                    "Loan Statement": _money_col(),
                    "Variance": _money_col(),
                },
            )
        else:
            st.caption("No escrow/reserve accounts in this trial balance matched a loan statement's escrow fields.")

        st.markdown("**Contributions / Distributions cross-check (vs. Distribution Waterfall)**")
        matched_tier = _match_ownership_tier(cfg, entity.entity_name)
        if not matched_tier:
            st.caption(
                f'No ownership tier is configured for "{entity.entity_name}" — nothing to cross-check at '
                "this level yet. This is expected for the property-level entity, which sits below the "
                "venture tier in the JV structure; it'll apply once a venture- or co-GP-level trial "
                "balance is uploaded."
            )
        elif not result or not result.waterfall or matched_tier.tier_id not in result.waterfall.tiers:
            st.caption(f'Matched to the "{matched_tier.label()}" tier, but no waterfall data is loaded to compare against.')
        else:
            wf_tier = result.waterfall.tiers[matched_tier.tier_id]
            wf_contrib = sum(inv.contributions_to_date or 0.0 for inv in wf_tier.investors)
            wf_dist = sum(abs(inv.distributions_to_date or 0.0) for inv in wf_tier.investors)
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "Metric": "Contributions",
                            "Trial Balance": entity.contributions,
                            "Waterfall (to date)": wf_contrib,
                            "Variance": (entity.contributions or 0.0) - wf_contrib,
                        },
                        {
                            "Metric": "Distributions",
                            "Trial Balance": entity.distributions,
                            "Waterfall (to date)": wf_dist,
                            "Variance": (entity.distributions or 0.0) - wf_dist,
                        },
                    ]
                ),
                width="stretch",
                hide_index=True,
                column_config={
                    "Trial Balance": _money_col(),
                    "Waterfall (to date)": _money_col(),
                    "Variance": _money_col(),
                },
            )
            st.caption(f'Matched to the "{matched_tier.label()}" tier ({matched_tier.distributing_entity}).')

        st.divider()


def _render_waterfall_tier(tier: WaterfallTier, indent: int = 0) -> None:
    prefix = "&nbsp;&nbsp;&nbsp;&nbsp;" * indent
    st.markdown(f"{prefix}**{tier.distributing_entity}**  ·  {tier.as_of_label}", unsafe_allow_html=True)

    render_kpi_tiles(
        [
            ("Net Cash Available", _fmt_money(tier.net_cash_available), None),
            ("Distribution Recommendation", _fmt_money(tier.distribution_recommendation), None),
            ("Cash Projected", _fmt_money(tier.cash_projected), None),
        ]
    )

    if tier.cash_holdbacks:
        with st.expander("Cash hold-backs"):
            for label, value in tier.cash_holdbacks.items():
                st.write(f"{label}: {_fmt_money(value)}")

    if tier.investors:
        df = pd.DataFrame(
            [
                {
                    "Investor": inv.display_name,
                    "Ownership %": _pct100(inv.ownership_pct),
                    "Distribution Amount": inv.distribution_amount,
                    "Contributions to Date": inv.contributions_to_date,
                    "Distributions to Date": inv.distributions_to_date,
                    "Net Capital After": inv.net_capital_after,
                }
                for inv in tier.investors
            ]
        )
        st.dataframe(
            df,
            width="stretch",
            hide_index=True,
            column_config={
                "Ownership %": _pct_col(),
                "Distribution Amount": _money_col(),
                "Contributions to Date": _money_col(),
                "Distributions to Date": _money_col(),
                "Net Capital After": _money_col(),
            },
        )


def _render_waterfall_tier_view(tier_cfg, result: Optional[DistributionWorkbookResult]) -> None:
    current_col, stabilized_col = st.columns(2)

    with current_col:
        st.markdown("##### Current Distribution")
        wf_tier = result.waterfall.tiers.get(tier_cfg.tier_id) if result and result.waterfall else None
        if wf_tier:
            _render_waterfall_tier(wf_tier)
        else:
            st.info("No distribution data available for this tier.")

    with stabilized_col:
        st.markdown("##### Stabilized Total")
        proj_tier = (
            result.projected_waterfall.tiers.get(tier_cfg.tier_id)
            if result and result.projected_waterfall
            else None
        )
        if proj_tier:
            st.write(proj_tier)
        else:
            st.info("Will populate once the leasing/investment model is finalized.")


def _render_waterfall(cfg: PropertyConfig, result: Optional[DistributionWorkbookResult]) -> None:
    # One tab per JV tier (e.g. LP/GP, Co-GP), so each can be reviewed independently,
    # plus a placeholder FF&GRC tier that isn't configured yet.
    tab_labels = [tier.label() for tier in cfg.ownership_tiers] + ["FF&GRC"]
    tabs = st.tabs(tab_labels)

    for tier_cfg, tab in zip(cfg.ownership_tiers, tabs):
        with tab:
            _render_waterfall_tier_view(tier_cfg, result)

    with tabs[-1]:
        st.info("FF&GRC — not yet configured.")

    if cfg.jv_documents:
        st.divider()
        st.markdown("#### JV Abstracts")
        for jv_ref in cfg.jv_documents:
            abstract = load_jv_abstract(cfg, jv_ref)
            with st.expander(jv_ref.name):
                if abstract:
                    st.json(abstract)
                else:
                    st.caption("No abstract on file yet — run `tools/extract_loan_abstract.py`.")


def _budget_lines_df(lines, label_fn) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Account": label_fn(line),
                "Budget": line.budget_value,
                "Actual": line.actual_value,
                "Variance $": line.variance_dollar,
                "Variance %": _pct100(line.variance_pct),
                "Missing": line.missing_side or "",
            }
            for line in lines
        ]
    )


def _budget_column_config():
    return {
        "Budget": _money_col(),
        "Actual": _money_col(),
        "Variance $": _money_col(),
        "Variance %": _pct_col(),
    }


def _render_budget(result: Optional[DistributionWorkbookResult]) -> None:
    if not result or not result.budget_comparison:
        st.info("No budget comparison data available.")
        return

    period = result.budget_comparison.period
    if period:
        st.caption(f"Period: {period.strftime('%B %Y')} (last month the Cash Flow tab labels as Actuals)")

    budget_subtabs = ["Summary", "Detailed"]
    if "budget_subtab" not in st.session_state or st.session_state.budget_subtab not in budget_subtabs:
        st.session_state.budget_subtab = budget_subtabs[0]
    sub = st.segmented_control("View", budget_subtabs, key="budget_subtab")
    if sub is None:
        st.session_state.budget_subtab = budget_subtabs[0]
        sub = budget_subtabs[0]

    if sub == "Summary":
        if not result.budget_summary or not result.budget_summary.lines:
            st.info("No P&L summary available.")
        else:
            st.dataframe(
                _budget_lines_df(result.budget_summary.lines, lambda line: line.account_label),
                width="stretch",
                hide_index=True,
                column_config=_budget_column_config(),
            )
            st.caption(
                "Expenses/NOI/Cash Flow after Debt and Capital are computed on the budget side "
                "(the Budget tab splits expenses into Recoverable/Non-Recoverable rather than one "
                "lump total) — same formula order the Cash Flow tab itself uses. Net Income comes "
                "from the distribution file's trailing-12-month Net Income row, actual only — that "
                "row has no forward/budget-scenario values, so there's no budget-side figure to show."
            )
    else:
        if not result.budget_comparison.lines:
            st.info("No account-level budget comparison available.")
        else:
            st.dataframe(
                _budget_lines_df(
                    result.budget_comparison.lines,
                    lambda line: f"{line.account_code} — {line.account_label}",
                ),
                width="stretch",
                hide_index=True,
                column_config=_budget_column_config(),
            )


def _render_annual_budget(result: Optional[DistributionWorkbookResult]) -> None:
    if not result or not result.annual_budget_summary or not result.annual_budget_summary.lines:
        st.info("No annual budget comparison available.")
        return

    ytd_through = result.annual_budget_summary.period
    if ytd_through:
        st.caption(
            f"Annual Budget = full-year {ytd_through.year} budget. YTD Actual = January "
            f"through {ytd_through.strftime('%B %Y')} (the last month the Cash Flow tab labels "
            "as Actuals)."
        )
    st.dataframe(
        _budget_lines_df(result.annual_budget_summary.lines, lambda line: line.account_label),
        width="stretch",
        hide_index=True,
        column_config=_budget_column_config(),
    )
    st.caption(
        "Capital Expenditures and Cash Flow after Debt Service aren't in the monthly Budget vs. "
        "Actuals view — added here to match the full-year waterfall Ryan's own Kardin budget "
        "report uses (Income → OpEx → NOI → Debt Service → CapEx). Net Income is left out: it's a "
        "trailing-12-month actual-only figure, a different concept from a January-through-YTD sum."
    )


def _render_holdsell_assumptions(
    cfg: PropertyConfig,
    result: Optional[DistributionWorkbookResult],
    rent_roll: Optional[RentRollResult],
    loan_statements: List[LoanStatement],
    entity_trial_balances: List[EntityTrialBalance],
    data_dir: str,
    opex_categories: Optional[dict] = None,
) -> None:
    opex_categories = opex_categories or {}
    st.caption(
        "Assumptions here drive the Hold/Sell Excel model — edit and save them below, then "
        "download a fresh copy of the workbook with these values and the latest real data baked "
        "in. All the actual math (rollover schedule, debt amortization, IRR) lives only in the "
        "workbook's own formulas, never duplicated here, so there's exactly one place it can be "
        "wrong instead of two that could disagree."
    )

    current = holdsell_model.load_assumptions(cfg, data_dir)

    sections: List[str] = []
    for _, _, section, _ in holdsell_model.ASSUMPTION_FIELDS:
        if section not in sections:
            sections.append(section)

    with st.form("holdsell_assumptions_form"):
        new_values: dict = {}
        for section in sections:
            st.markdown(f"**{section}**")
            for key, label, field_section, kind in holdsell_model.ASSUMPTION_FIELDS:
                if field_section != section:
                    continue
                stored = current.get(key)
                if kind == "pct":
                    shown = stored * 100 if stored is not None else None
                    val = st.number_input(f"{label}", value=shown, step=0.1, format="%.2f", key=f"hs_{key}")
                    new_values[key] = (val / 100) if val is not None else None
                elif kind in ("years", "months"):
                    val = st.number_input(f"{label}", value=stored, step=1.0, format="%.0f", key=f"hs_{key}")
                    new_values[key] = val
                elif kind == "dollar_psf":
                    val = st.number_input(f"{label}", value=stored, step=0.5, format="%.2f", key=f"hs_{key}")
                    new_values[key] = val
                else:  # dollar
                    val = st.number_input(f"{label}", value=stored, step=1000.0, format="%.0f", key=f"hs_{key}")
                    new_values[key] = val

        st.markdown("**SOFR Curve (Monthly)**")
        st.caption(
            "Optional — leave a month blank to fall back to the derived Current SOFR. Paste in a "
            "real forward curve here if you have one; it round-trips into every future export."
        )
        stored_curve = current.get("sofr_curve") or [None] * holdsell_model.N_MONTHS
        curve_df = pd.DataFrame(
            {
                "Month": list(range(1, holdsell_model.N_MONTHS + 1)),
                "SOFR %": [v * 100 if v is not None else None for v in stored_curve],
            }
        )
        edited_curve = st.data_editor(
            curve_df,
            key="hs_sofr_curve",
            hide_index=True,
            disabled=["Month"],
            column_config={"SOFR %": st.column_config.NumberColumn(format="%.2f", step=0.01)},
            height=200,
        )

        saved = st.form_submit_button("Save Assumptions")
        if saved:
            new_values["sofr_curve"] = [
                (v / 100) if v is not None and not pd.isna(v) else None for v in edited_curve["SOFR %"]
            ]
            holdsell_model.save_assumptions(cfg, new_values, data_dir)
            st.success("Assumptions saved.")
            current = new_values

    st.divider()
    if not result or not rent_roll:
        st.info("Download isn't available yet — the Hold/Sell model needs a distribution workbook and rent roll loaded for this period.")
        return

    try:
        wb = holdsell_model.build_workbook(
            cfg, result, rent_roll, loan_statements, entity_trial_balances, current, opex_categories
        )
        buf = io.BytesIO()
        wb.save(buf)
        st.download_button(
            "Download Excel Model",
            data=buf.getvalue(),
            file_name=f"{cfg.display()} - Hold-Sell Model.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Couldn't build the export: {exc}")
