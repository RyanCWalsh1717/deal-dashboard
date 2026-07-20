"""Deal Dashboard — internal Streamlit app for tracking active/owned CRE deals.

Distinct from the `acquisitions` project (screens incoming OMs pre-acquisition).
"""

import os
import sys
from pathlib import Path
from typing import Optional, Tuple

import streamlit as st

APP_DIR = Path(__file__).parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from pipeline import source_files
from pipeline.models import PortfolioSummaryRow
from pipeline.parsers.cash_accounts import parse_loan_statement, parse_trial_balance_cash_accounts
from pipeline.parsers.distribution_workbook import parse_workbook
from pipeline.parsers.rent_roll import parse_rent_roll
from pipeline.property_config import PropertyConfig, discover_properties
from views.branding import render_hero
from views.portfolio_view import render_portfolio
from views.property_detail_view import render_property_detail

DATA_DIR = APP_DIR / "data"

st.set_page_config(page_title="Deal Dashboard", layout="wide", initial_sidebar_state="expanded")


def check_password() -> bool:
    if st.session_state.get("authenticated"):
        return True
    render_hero("Deal Dashboard", "Greatland Realty Partners &mdash; Active Deal Tracking")
    st.write("")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("#### Sign In")
        password = st.text_input("Password", type="password", key="password_input")
        if st.button("Sign In", width="stretch"):
            correct = st.secrets.get("APP_PASSWORD") or os.environ.get("APP_PASSWORD", "")
            if password == correct and correct != "":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    return False


@st.cache_data(show_spinner="Parsing distribution workbook...")
def _cached_parse(path_str: str, mtime: float, property_code: str, data_dir: str):
    cfg = PropertyConfig.load(property_code, data_dir=data_dir)
    return parse_workbook(path_str, cfg)


def _resolve_workbook_path(cfg: PropertyConfig, period: Optional[str]) -> Tuple[Optional[Path], Optional[float]]:
    if not period:
        return None, None
    hint_name = Path(cfg.source_workbook_path_hint).name
    p = source_files.resolve_period_file(cfg, period, hint_name, str(DATA_DIR))
    return (p, p.stat().st_mtime) if p else (None, None)


@st.cache_data(show_spinner="Parsing trial balance...")
def _cached_trial_balance(path_str: str, mtime: float, yardi_codes: tuple):
    return parse_trial_balance_cash_accounts(path_str, list(yardi_codes))


def _resolve_trial_balance_path(cfg: PropertyConfig, period: Optional[str]) -> Tuple[Optional[Path], Optional[float]]:
    if not period:
        return None, None
    p = source_files.resolve_period_file(cfg, period, "trial_balance.xlsx", str(DATA_DIR))
    return (p, p.stat().st_mtime) if p else (None, None)


@st.cache_data(show_spinner="Parsing loan statement...")
def _cached_loan_statement(path_str: str, mtime: float):
    return parse_loan_statement(path_str)


def _discover_loan_statement_paths(cfg: PropertyConfig, period: Optional[str]) -> dict:
    """Returns {path_str: mtime} for every loan-statement PDF resolved for
    this property + period (with carry-forward to the nearest earlier
    period that has any, via `source_files.resolve_period_loan_statements`)."""
    if not period:
        return {}
    return {str(p): p.stat().st_mtime for p in source_files.resolve_period_loan_statements(cfg, period, str(DATA_DIR))}


@st.cache_data(show_spinner="Parsing rent roll...")
def _cached_rent_roll(path_str: str, mtime: float, property_code: str):
    import openpyxl

    wb = openpyxl.load_workbook(path_str, data_only=True)
    return parse_rent_roll(wb["Report1"] if "Report1" in wb.sheetnames else wb.worksheets[0], property_code)


def _resolve_rent_roll_path(cfg: PropertyConfig, period: Optional[str]) -> Tuple[Optional[Path], Optional[float]]:
    if not period:
        return None, None
    p = source_files.resolve_period_file(cfg, period, "rent_roll.xlsx", str(DATA_DIR))
    return (p, p.stat().st_mtime) if p else (None, None)


def _build_portfolio_row(cfg: PropertyConfig, result) -> PortfolioSummaryRow:
    total_cash = None
    total_debt = None
    last_dist_amount = None
    last_dist_as_of = None

    if result:
        cash_values = [pos.total_cash for pos in result.equity.values() if pos.total_cash is not None]
        total_cash = sum(cash_values) if cash_values else None
        total_debt = result.debt.total_outstanding if result.debt else None
        if result.waterfall:
            top_tier_cfg = cfg.top_level_tier()
            top = result.waterfall.tiers.get(top_tier_cfg.tier_id) if top_tier_cfg else None
            if top is None and result.waterfall.tiers:
                top = next(iter(result.waterfall.tiers.values()))
            if top:
                last_dist_amount = top.distribution_recommendation
                last_dist_as_of = top.as_of_label

    return PortfolioSummaryRow(
        property_code=cfg.property_code,
        display_name=cfg.display(),
        address=cfg.property_address,
        market=cfg.market,
        investor_names=cfg.investor_display_names(),
        total_cash=total_cash,
        total_debt_outstanding=total_debt,
        last_distribution_amount=last_dist_amount,
        last_distribution_as_of=last_dist_as_of,
    )


def _process_uploads(cfg: PropertyConfig, uploaded_files) -> Optional[str]:
    """Classifies + saves every file in a batch upload; returns the newest
    period touched (so the picker can jump straight to it), or None if
    nothing was successfully saved."""
    newest_period = None
    for f in uploaded_files:
        data = f.getvalue()
        classified = source_files.classify_upload(data, f.name, cfg)
        if classified.file_type == "unknown" or not classified.period:
            st.sidebar.warning(classified.error or f"Couldn't process {f.name}.")
            continue
        source_files.save_classified_upload(classified, data, cfg, str(DATA_DIR))
        st.sidebar.success(f"{f.name} → {classified.file_type.replace('_', ' ')} ({classified.period})")
        if newest_period is None or classified.period > newest_period:
            newest_period = classified.period
    return newest_period


def main() -> None:
    for key, default in {
        "selected_property": None,
        "upload_epoch": {},
    }.items():
        if key not in st.session_state:
            st.session_state[key] = default

    properties = discover_properties(str(DATA_DIR))

    st.sidebar.markdown(
        "<h2 style='color:#1A5C22;'>Deal Dashboard</h2>", unsafe_allow_html=True
    )
    view = st.sidebar.radio("View", ["Portfolio", "Property Detail"])

    if not properties:
        st.warning(
            "No properties configured yet. Copy `data/TEMPLATE/config.yaml` to "
            "`data/<property_code>/config.yaml` to add one."
        )
        return

    if view == "Portfolio":
        rows = []
        for cfg in properties:
            periods = source_files.list_periods(cfg, str(DATA_DIR))
            latest_period = periods[0] if periods else None
            path, mtime = _resolve_workbook_path(cfg, latest_period)
            result = _cached_parse(str(path), mtime, cfg.property_code, str(DATA_DIR)) if path else None
            rows.append(_build_portfolio_row(cfg, result))
        render_portfolio(rows)
        return

    codes = [c.property_code for c in properties]
    names = {c.property_code: c.display() for c in properties}
    selected_code = st.sidebar.selectbox("Property", codes, format_func=lambda c: names[c])
    cfg = next(c for c in properties if c.property_code == selected_code)

    period_key = f"period_select_{selected_code}"
    if "pending_period" in st.session_state:
        st.session_state[period_key] = st.session_state.pop("pending_period")

    periods = source_files.list_periods(cfg, str(DATA_DIR))
    selected_period = None
    if periods:
        if period_key not in st.session_state:
            st.session_state[period_key] = periods[0]
        selected_period = st.sidebar.selectbox("Viewing Period", periods, key=period_key)

    st.sidebar.markdown("---")
    with st.sidebar.expander("Update Source Files"):
        st.caption(
            "Drop in an updated distribution workbook, trial balance, rent roll, "
            "or loan statement PDF(s) — each file is auto-detected by its "
            "content and filed under the period it's dated as of."
        )
        epoch = st.session_state.upload_epoch.get(selected_code, 0)
        uploaded_files = st.file_uploader(
            "Files",
            type=["xlsx", "pdf"],
            accept_multiple_files=True,
            key=f"multi_upload_{selected_code}_{epoch}",
            label_visibility="collapsed",
        )
        if uploaded_files:
            newest_period = _process_uploads(cfg, uploaded_files)
            if newest_period:
                st.session_state.pending_period = newest_period
            st.session_state.upload_epoch[selected_code] = epoch + 1
            st.rerun()

    if not periods:
        st.info(
            f"No source files yet for **{cfg.display()}**. Use “Update Source Files” "
            "in the sidebar to upload the distribution workbook, trial balance, rent roll, "
            "and/or loan statements."
        )
        return

    path, mtime = _resolve_workbook_path(cfg, selected_period)
    result = None
    if path is None:
        st.info(f"No distribution workbook found for **{cfg.display()}** as of {selected_period}.")
    else:
        try:
            result = _cached_parse(str(path), mtime, cfg.property_code, str(DATA_DIR))
        except Exception as exc:
            st.error(f"Failed to parse workbook: {exc}")

    tb_path, tb_mtime = _resolve_trial_balance_path(cfg, selected_period)
    cash_accounts = []
    if tb_path is not None:
        try:
            cash_accounts = _cached_trial_balance(str(tb_path), tb_mtime, tuple(cfg.yardi_codes))
        except Exception as exc:
            st.error(f"Failed to parse trial balance: {exc}")

    loan_statements = []
    for path_str, ls_mtime in _discover_loan_statement_paths(cfg, selected_period).items():
        try:
            stmt = _cached_loan_statement(path_str, ls_mtime)
            if stmt:
                loan_statements.append(stmt)
        except Exception as exc:
            st.error(f"Failed to parse loan statement ({Path(path_str).name}): {exc}")

    if not cash_accounts and loan_statements:
        from pipeline.parsers.cash_accounts import loan_statement_cash_accounts

        for stmt in loan_statements:
            cash_accounts.extend(loan_statement_cash_accounts(stmt))

    rr_path, rr_mtime = _resolve_rent_roll_path(cfg, selected_period)
    rent_roll = None
    if rr_path is not None:
        try:
            rent_roll = _cached_rent_roll(str(rr_path), rr_mtime, cfg.property_code)
        except Exception as exc:
            st.error(f"Failed to parse rent roll: {exc}")

    render_property_detail(cfg, result, cash_accounts, rent_roll, loan_statements)


if not check_password():
    st.stop()

main()
