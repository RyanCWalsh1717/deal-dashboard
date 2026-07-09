"""Portfolio view — one row per property, filterable by market and investor."""

from __future__ import annotations

from typing import List

import pandas as pd
import streamlit as st

from pipeline.models import PortfolioSummaryRow
from views.branding import render_hero


def _fmt_money(v):
    return "—" if v is None else f"${v:,.0f}"


def render_portfolio(rows: List[PortfolioSummaryRow]) -> None:
    render_hero("Portfolio", "Greatland Realty Partners &mdash; All Properties")
    st.write("")

    if not rows:
        st.info("No properties to show.")
        return

    markets = sorted({r.market for r in rows if r.market})
    all_investors = sorted({name for r in rows for name in r.investor_names})

    col1, col2 = st.columns(2)
    with col1:
        selected_markets = st.multiselect("Market", markets)
    with col2:
        selected_investors = st.multiselect("Investor", all_investors)

    filtered = [
        r
        for r in rows
        if (not selected_markets or r.market in selected_markets)
        and (not selected_investors or any(i in selected_investors for i in r.investor_names))
    ]

    df = pd.DataFrame(
        [
            {
                "Property": r.display_name,
                "Market": r.market,
                "Address": r.address,
                "Investors": ", ".join(r.investor_names),
                "Total Cash": _fmt_money(r.total_cash),
                "Total Debt Outstanding": _fmt_money(r.total_debt_outstanding),
                "Last Distribution": _fmt_money(r.last_distribution_amount),
                "As Of": r.last_distribution_as_of or "—",
            }
            for r in filtered
        ]
    )

    if df.empty:
        st.warning("No properties match the selected filters.")
    else:
        st.dataframe(df, width="stretch", hide_index=True)
