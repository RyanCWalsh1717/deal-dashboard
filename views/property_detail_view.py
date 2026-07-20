"""Property detail view: Summary / Cash / Equity & Capital / Debt & Loans /
Distribution Waterfall (per-tier tabs, each with Current + Stabilized) /
Budget vs. Actuals (Summary + Detailed) / Sources & Uses (placeholder) /
Leasing & Investment Outlook (placeholder).
"""

from __future__ import annotations

import re
from typing import List, Optional

import pandas as pd
import streamlit as st

from pipeline.models import (
    CashAccountBalance,
    DistributionWorkbookResult,
    LoanStatement,
    RentRollResult,
    WaterfallTier,
)
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


SECTIONS = [
    "Summary",
    "Cash",
    "Equity & Capital",
    "Balance Sheet",
    "Debt & Loans",
    "Distribution Waterfall",
    "Budget vs. Actuals",
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
) -> None:
    cash_accounts = cash_accounts or []
    loan_statements = loan_statements or []

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
        _render_summary(cfg, result, rent_roll)
    elif section == "Cash":
        _render_cash(result, cash_accounts)
    elif section == "Equity & Capital":
        _render_equity(result)
    elif section == "Balance Sheet":
        _render_balance_sheet(result)
    elif section == "Debt & Loans":
        _render_debt(cfg, result, loan_statements)
    elif section == "Distribution Waterfall":
        _render_waterfall(cfg, result)
    elif section == "Budget vs. Actuals":
        _render_budget(result)
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
    cfg: PropertyConfig, result: Optional[DistributionWorkbookResult], rent_roll: Optional[RentRollResult]
) -> None:
    cash_on_hand = _total_cash(result)
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
            ("Cash on Hand", _fmt_money(cash_on_hand), None),
            ("NOI (Last Actual Month)", _fmt_money(noi_today), noi_sub),
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
    if noi_period:
        st.caption(f"NOI as of {noi_period.strftime('%B %Y')} — the last month the Cash Flow tab labels as Actuals.")
    st.caption("Stabilized NOI will populate once the leasing/investment outlook model is finalized.")

    st.divider()
    st.markdown("#### Loan Terms")
    if result and result.debt and result.debt.tranches:
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
        st.caption("Full loan terms (maturity, extension options, covenants) pending loan abstracts.")
        if st.button("View Debt & Loans →", key="goto_debt"):
            _goto("Debt & Loans")
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

        df = pd.DataFrame(
            [
                {
                    "Unit": line.unit_code,
                    "Floor": line.floor,
                    "Tenant": line.tenant_name if not line.is_vacant else "VACANT",
                    "SF": line.unit_area,
                    "Lease Start": line.lease_from,
                    "Lease End": line.lease_to,
                    "Annual Rent": line.annual_rent,
                    "Rent/SF": line.annual_rent_psf,
                    "Lease Type": line.lease_type,
                }
                for line in rent_roll.lines
            ]
        )
        st.dataframe(
            df,
            width="stretch",
            hide_index=True,
            column_config={
                "SF": st.column_config.NumberColumn(format="%,d"),
                "Annual Rent": _money_col(),
                "Rent/SF": st.column_config.NumberColumn(format="$%,.2f"),
            },
        )


def _render_cash(result: Optional[DistributionWorkbookResult], cash_accounts: List[CashAccountBalance]) -> None:
    boxes = [("Operating Cash (all entities)", _total_cash(result), "equity_tabs")]
    boxes.append(("DACA", None, "placeholder"))

    if cash_accounts:
        seen = set()
        for acct in cash_accounts:
            if acct.label in seen:
                continue
            seen.add(acct.label)
            boxes.append((acct.label, acct.balance, acct.source))

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
