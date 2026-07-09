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
from views.branding import render_hero


def _fmt_money(v):
    return "—" if v is None else f"${v:,.0f}"


def _fmt_pct(v):
    return "—" if v is None else f"{v * 100:.1f}%"


def _total_cash(result: Optional[DistributionWorkbookResult]) -> Optional[float]:
    if not result or not result.equity:
        return None
    values = [pos.total_cash for pos in result.equity.values() if pos.total_cash is not None]
    return sum(values) if values else None


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

    tabs = st.tabs(
        [
            "Summary",
            "Cash",
            "Equity & Capital",
            "Debt & Loans",
            "Distribution Waterfall",
            "Budget vs. Actuals",
            "Sources & Uses",
            "Leasing & Investment Outlook",
        ]
    )

    with tabs[0]:
        _render_summary(cfg, result, rent_roll)
    with tabs[1]:
        _render_cash(result, cash_accounts)
    with tabs[2]:
        _render_equity(result)
    with tabs[3]:
        _render_debt(cfg, result, loan_statements)
    with tabs[4]:
        _render_waterfall(cfg, result)
    with tabs[5]:
        _render_budget(result)
    with tabs[6]:
        st.info(
            "Sources & Uses will populate once the leasing/investment outlook model is finalized — "
            "it lives entirely in that model, not the distribution workbook."
        )
    with tabs[7]:
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

    col1, col2, col3 = st.columns(3)
    col1.metric("Cash on Hand", _fmt_money(cash_on_hand))
    col2.metric("NOI (Last Actual Month)", _fmt_money(noi_today))
    col3.metric("Stabilized NOI", "—")
    if noi_period:
        st.caption(f"NOI as of {noi_period.strftime('%B %Y')} — the last month the Cash Flow tab labels as Actuals.")
    st.caption("Stabilized NOI will populate once the leasing/investment outlook model is finalized.")

    st.divider()
    st.markdown("#### Loan Terms")
    if result and result.debt and result.debt.tranches:
        df = pd.DataFrame(
            [
                {"Tranche": t.tranche_name, "Balance": _fmt_money(t.outstanding_balance), "Rate": _fmt_pct(t.interest_rate)}
                for t in result.debt.tranches
            ]
        )
        st.dataframe(df, width="stretch", hide_index=True)
        st.caption("Full loan terms (maturity, extension options, covenants) pending loan abstracts.")
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
        rr1, rr2, rr3 = st.columns(3)
        rr1.metric("Occupancy", _fmt_pct(rent_roll.occupancy_pct))
        rr2.metric("Leased SF", f"{rent_roll.total_leased_sf:,.0f}")
        rr3.metric("Vacant SF", f"{rent_roll.total_vacant_sf:,.0f}")

        df = pd.DataFrame(
            [
                {
                    "Unit": line.unit_code,
                    "Floor": line.floor,
                    "Tenant": line.tenant_name if not line.is_vacant else "VACANT",
                    "SF": line.unit_area,
                    "Lease Start": line.lease_from,
                    "Lease End": line.lease_to,
                    "Annual Rent": _fmt_money(line.annual_rent),
                    "Rent/SF": f"${line.annual_rent_psf:,.2f}" if line.annual_rent_psf else "—",
                    "Lease Type": line.lease_type,
                }
                for line in rent_roll.lines
            ]
        )
        st.dataframe(df, width="stretch", hide_index=True)


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

    cols = st.columns(4)
    for i, (label, value, _source) in enumerate(boxes):
        with cols[i % 4]:
            st.metric(label, _fmt_money(value))

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
    st.metric("Latest Month Net Activity", _fmt_money(latest_total))

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
        col1, col2, col3 = st.columns(3)
        col1.metric("Cash Balance", _fmt_money(pos.total_cash))
        col2.metric("Total Contributions", _fmt_money(pos.total_contributions))
        col3.metric("Total Distributions", _fmt_money(pos.total_distributions))

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
            st.dataframe(df, width="stretch", hide_index=True)
        st.divider()

    st.markdown("#### Future Distributions")
    st.info("Will populate once the leasing/investment outlook model is finalized.")


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
                    "Outstanding Balance": _fmt_money(t.outstanding_balance),
                    "Spread over SOFR": _fmt_pct(t.interest_rate),
                }
                for t in debt.tranches
            ]
        )
        st.dataframe(df, width="stretch", hide_index=True)
        st.metric("Total Outstanding (Forecast)", _fmt_money(debt.total_outstanding))
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
                    "Actual Balance": _fmt_money(stmt.principal_balance),
                    "Actual Rate (all-in)": _fmt_pct(stmt.interest_rate),
                    "vs. Forecast Balance": _fmt_money(balance_delta) if balance_delta is not None else "—",
                }
            )
        st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
        total_actual = sum(s.principal_balance or 0 for s in loan_statements)
        st.metric("Total Outstanding (Actual)", _fmt_money(total_actual))

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

    cols = st.columns(3)
    cols[0].metric("Net Cash Available", _fmt_money(tier.net_cash_available))
    cols[1].metric("Distribution Recommendation", _fmt_money(tier.distribution_recommendation))
    cols[2].metric("Cash Projected", _fmt_money(tier.cash_projected))

    if tier.cash_holdbacks:
        with st.expander("Cash hold-backs"):
            for label, value in tier.cash_holdbacks.items():
                st.write(f"{label}: {_fmt_money(value)}")

    if tier.investors:
        df = pd.DataFrame(
            [
                {
                    "Investor": inv.display_name,
                    "Ownership %": _fmt_pct(inv.ownership_pct),
                    "Distribution Amount": _fmt_money(inv.distribution_amount),
                    "Contributions to Date": _fmt_money(inv.contributions_to_date),
                    "Distributions to Date": _fmt_money(inv.distributions_to_date),
                    "Net Capital After": _fmt_money(inv.net_capital_after),
                }
                for inv in tier.investors
            ]
        )
        st.dataframe(df, width="stretch", hide_index=True)


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
                "Variance %": _fmt_pct(line.variance_pct),
                "Missing": line.missing_side or "",
            }
            for line in lines
        ]
    )


def _render_budget(result: Optional[DistributionWorkbookResult]) -> None:
    if not result or not result.budget_comparison:
        st.info("No budget comparison data available.")
        return

    period = result.budget_comparison.period
    if period:
        st.caption(f"Period: {period.strftime('%B %Y')} (last month the Cash Flow tab labels as Actuals)")

    summary_tab, detailed_tab = st.tabs(["Summary", "Detailed"])

    with summary_tab:
        if not result.budget_summary or not result.budget_summary.lines:
            st.info("No P&L summary available.")
        else:
            st.dataframe(
                _budget_lines_df(result.budget_summary.lines, lambda line: line.account_label),
                width="stretch",
                hide_index=True,
            )
            st.caption(
                "Expenses/NOI/Cash Flow after Debt and Capital are computed on the budget side "
                "(the Budget tab splits expenses into Recoverable/Non-Recoverable rather than one "
                "lump total) — same formula order the Cash Flow tab itself uses. Net Income comes "
                "from the distribution file's trailing-12-month Net Income row, actual only — that "
                "row has no forward/budget-scenario values, so there's no budget-side figure to show."
            )

    with detailed_tab:
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
            )
