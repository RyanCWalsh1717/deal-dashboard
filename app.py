"""Deal Dashboard — internal Streamlit app for tracking active/owned CRE deals.

Distinct from the `acquisitions` project (screens incoming OMs pre-acquisition).
"""

import os
import sys
import tempfile
from pathlib import Path
from typing import Optional, Tuple

import streamlit as st

APP_DIR = Path(__file__).parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from pipeline.models import PortfolioSummaryRow
from pipeline.parsers.cash_accounts import (
    loan_statement_cash_accounts,
    parse_loan_statement,
    parse_trial_balance_cash_accounts,
)
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


def _resolve_workbook_path(cfg: PropertyConfig) -> Tuple[Optional[Path], Optional[float]]:
    override = st.session_state.workbook_override_path.get(cfg.property_code)
    if override:
        p = Path(override)
        if p.exists():
            return p, p.stat().st_mtime
    p = cfg.workbook_path(str(DATA_DIR))
    if p.exists():
        return p, p.stat().st_mtime
    return None, None


@st.cache_data(show_spinner="Parsing trial balance...")
def _cached_trial_balance(path_str: str, mtime: float, yardi_codes: tuple):
    return parse_trial_balance_cash_accounts(path_str, list(yardi_codes))


def _resolve_trial_balance_path(cfg: PropertyConfig) -> Tuple[Optional[Path], Optional[float]]:
    override = st.session_state.trial_balance_override_path.get(cfg.property_code)
    if override:
        p = Path(override)
        if p.exists():
            return p, p.stat().st_mtime
    p = DATA_DIR / cfg.property_code / "source_files" / "trial_balance.xlsx"
    if p.exists():
        return p, p.stat().st_mtime
    return None, None


@st.cache_data(show_spinner="Parsing loan statement...")
def _cached_loan_statement(path_str: str, mtime: float):
    return parse_loan_statement(path_str)


def _discover_loan_statement_paths(cfg: PropertyConfig) -> dict:
    """Returns {path_str: mtime} for every loan-statement PDF known for this
    property — local-convention files plus any sidebar-uploaded overrides."""
    paths = {}
    base = DATA_DIR / cfg.property_code / "source_files" / "loan_statements"
    if base.exists():
        for p in sorted(base.glob("*.pdf")):
            paths[str(p)] = p.stat().st_mtime
    for override in st.session_state.loan_statement_override_paths.get(cfg.property_code, []):
        p = Path(override)
        if p.exists():
            paths[str(p)] = p.stat().st_mtime
    return paths


@st.cache_data(show_spinner="Parsing rent roll...")
def _cached_rent_roll(path_str: str, mtime: float, property_code: str):
    import openpyxl

    wb = openpyxl.load_workbook(path_str, data_only=True)
    return parse_rent_roll(wb["Report1"] if "Report1" in wb.sheetnames else wb.worksheets[0], property_code)


def _resolve_rent_roll_path(cfg: PropertyConfig) -> Tuple[Optional[Path], Optional[float]]:
    override = st.session_state.rent_roll_override_path.get(cfg.property_code)
    if override:
        p = Path(override)
        if p.exists():
            return p, p.stat().st_mtime
    p = DATA_DIR / cfg.property_code / "source_files" / "rent_roll.xlsx"
    if p.exists():
        return p, p.stat().st_mtime
    return None, None


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


def main() -> None:
    for key, default in {
        "selected_property": None,
        "workbook_override_path": {},
        "trial_balance_override_path": {},
        "loan_statement_override_paths": {},
        "rent_roll_override_path": {},
    }.items():
        if key not in st.session_state:
            st.session_state[key] = default

    properties = discover_properties(str(DATA_DIR))

    st.sidebar.markdown(
        "<h2 style='color:#1F3864;'>Deal Dashboard</h2>", unsafe_allow_html=True
    )
    view = st.sidebar.radio("View", ["Portfolio", "Property Detail", "Style Preview"])

    if view == "Style Preview":
        from views.style_preview import render_style_preview

        render_style_preview()
        return

    if not properties:
        st.warning(
            "No properties configured yet. Copy `data/TEMPLATE/config.yaml` to "
            "`data/<property_code>/config.yaml` to add one."
        )
        return

    if view == "Portfolio":
        rows = []
        for cfg in properties:
            path, mtime = _resolve_workbook_path(cfg)
            result = _cached_parse(str(path), mtime, cfg.property_code, str(DATA_DIR)) if path else None
            rows.append(_build_portfolio_row(cfg, result))
        render_portfolio(rows)
        return

    codes = [c.property_code for c in properties]
    names = {c.property_code: c.display() for c in properties}
    selected_code = st.sidebar.selectbox("Property", codes, format_func=lambda c: names[c])
    cfg = next(c for c in properties if c.property_code == selected_code)

    st.sidebar.markdown("---")
    uploaded = st.sidebar.file_uploader(
        "Override workbook (.xlsx)", type=["xlsx"], key=f"upload_{selected_code}"
    )
    if uploaded is not None:
        tmp_dir = Path(tempfile.gettempdir()) / "deal_dashboard_uploads"
        tmp_dir.mkdir(exist_ok=True)
        tmp_path = tmp_dir / f"{selected_code}_{uploaded.name}"
        tmp_path.write_bytes(uploaded.getbuffer())
        st.session_state.workbook_override_path[selected_code] = str(tmp_path)
        st.rerun()

    uploaded_tb = st.sidebar.file_uploader(
        "Trial Balance (.xlsx)", type=["xlsx"], key=f"upload_tb_{selected_code}"
    )
    if uploaded_tb is not None:
        tmp_dir = Path(tempfile.gettempdir()) / "deal_dashboard_uploads"
        tmp_dir.mkdir(exist_ok=True)
        tmp_path = tmp_dir / f"{selected_code}_tb_{uploaded_tb.name}"
        tmp_path.write_bytes(uploaded_tb.getbuffer())
        st.session_state.trial_balance_override_path[selected_code] = str(tmp_path)
        st.rerun()

    uploaded_loans = st.sidebar.file_uploader(
        "Loan Statements (Berkadia .pdf, one per tranche)",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"upload_loans_{selected_code}",
    )
    if uploaded_loans:
        tmp_dir = Path(tempfile.gettempdir()) / "deal_dashboard_uploads"
        tmp_dir.mkdir(exist_ok=True)
        new_paths = []
        for f in uploaded_loans:
            tmp_path = tmp_dir / f"{selected_code}_loan_{f.name}"
            tmp_path.write_bytes(f.getbuffer())
            new_paths.append(str(tmp_path))
        st.session_state.loan_statement_override_paths[selected_code] = new_paths
        st.rerun()

    path, mtime = _resolve_workbook_path(cfg)
    result = None
    if path is None:
        st.info(
            f"No workbook found for **{cfg.display()}**. Drop one at "
            f"`{cfg.workbook_path(str(DATA_DIR))}` or upload one in the sidebar."
        )
    else:
        try:
            result = _cached_parse(str(path), mtime, cfg.property_code, str(DATA_DIR))
        except Exception as exc:
            st.error(f"Failed to parse workbook: {exc}")

    tb_path, tb_mtime = _resolve_trial_balance_path(cfg)
    cash_accounts = []
    if tb_path is not None:
        try:
            cash_accounts = _cached_trial_balance(str(tb_path), tb_mtime, tuple(cfg.yardi_codes))
        except Exception as exc:
            st.error(f"Failed to parse trial balance: {exc}")

    loan_statements = []
    for path_str, ls_mtime in _discover_loan_statement_paths(cfg).items():
        try:
            stmt = _cached_loan_statement(path_str, ls_mtime)
            if stmt:
                loan_statements.append(stmt)
        except Exception as exc:
            st.error(f"Failed to parse loan statement ({Path(path_str).name}): {exc}")

    if not cash_accounts and loan_statements:
        for stmt in loan_statements:
            cash_accounts.extend(loan_statement_cash_accounts(stmt))

    uploaded_rent_roll = st.sidebar.file_uploader(
        "Rent Roll source (Tenancy Schedule .xlsx)",
        type=["xlsx"],
        key=f"upload_rent_roll_{selected_code}",
    )
    if uploaded_rent_roll is not None:
        tmp_dir = Path(tempfile.gettempdir()) / "deal_dashboard_uploads"
        tmp_dir.mkdir(exist_ok=True)
        tmp_path = tmp_dir / f"{selected_code}_rentroll_{uploaded_rent_roll.name}"
        tmp_path.write_bytes(uploaded_rent_roll.getbuffer())
        st.session_state.rent_roll_override_path[selected_code] = str(tmp_path)
        st.rerun()

    rr_path, rr_mtime = _resolve_rent_roll_path(cfg)
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
