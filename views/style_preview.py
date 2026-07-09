"""Temporary side-by-side comparison of 3 metric-box styling directions —
not part of the real app, just for Ryan to pick a direction from. Remove (and
drop the "Style Preview" radio option in app.py) once a style is chosen.
"""

from __future__ import annotations

import streamlit as st

_SAMPLE_METRICS = [
    ("Cash on Hand", "$9,580", None),
    ("NOI (Last Actual Month)", "$814,214", "as of May 2026"),
    ("Stabilized NOI", "—", "pending leasing model"),
    ("Total Debt Outstanding", "$87,368,303", "forecast"),
    ("Occupancy", "84.6%", "148,754 SF leased"),
]


def render_style_preview() -> None:
    st.markdown("## Style Preview")
    st.caption("Comparing 3 metric-box directions with the same sample data — pick one and I'll apply it everywhere.")

    st.divider()
    st.markdown("### 1 — Clean Bordered Cards")
    st.caption("Simple st.container(border=True) boxes, minimal and professional.")
    _render_clean_cards()

    st.divider()
    st.markdown("### 2 — GRP-Branded KPI Tiles")
    st.caption("Colored tiles matching the green hero banner, more visually rich.")
    _render_grp_tiles()

    st.divider()
    st.markdown("### 3 — Dense Financial Dashboard")
    st.caption("Tighter spacing, bordered table rows, optimized for more numbers at once.")
    _render_dense_dashboard()


def _render_clean_cards() -> None:
    cols = st.columns(len(_SAMPLE_METRICS))
    for col, (label, value, sub) in zip(cols, _SAMPLE_METRICS):
        with col:
            with st.container(border=True):
                st.caption(label)
                st.markdown(f"### {value}")
                if sub:
                    st.caption(sub)


def _render_grp_tiles() -> None:
    tiles_html = ""
    for label, value, sub in _SAMPLE_METRICS:
        sub_html = f'<div class="grp-tile-sub">{sub}</div>' if sub else ""
        tiles_html += f"""
        <div class="grp-tile">
          <div class="grp-tile-label">{label}</div>
          <div class="grp-tile-value">{value}</div>
          {sub_html}
        </div>"""

    st.markdown(
        f"""
        <style>
          .grp-tile-row {{ display: flex; gap: 12px; flex-wrap: wrap; }}
          .grp-tile {{
            background: #E8F5E9;
            border: 1px solid #A5D6A7;
            border-left: 4px solid #1A5C22;
            border-radius: 8px;
            padding: 14px 18px;
            flex: 1;
            min-width: 150px;
          }}
          .grp-tile-label {{
            font-size: 0.78rem;
            color: #2E7D32;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.03em;
            margin-bottom: 4px;
          }}
          .grp-tile-value {{
            font-size: 1.5rem;
            font-weight: 700;
            color: #1A5C22;
            line-height: 1.2;
          }}
          .grp-tile-sub {{
            font-size: 0.72rem;
            color: #558B2F;
            margin-top: 4px;
          }}
        </style>
        <div class="grp-tile-row">{tiles_html}</div>
        """,
        unsafe_allow_html=True,
    )


def _render_dense_dashboard() -> None:
    rows_html = ""
    for i, (label, value, sub) in enumerate(_SAMPLE_METRICS):
        bg = "#F9FAFB" if i % 2 == 0 else "#FFFFFF"
        sub_html = f'<span style="color:#9E9E9E; font-size:0.7rem;"> &nbsp;{sub}</span>' if sub else ""
        rows_html += f"""
        <tr style="background:{bg};">
          <td style="padding:6px 12px; font-size:0.8rem; color:#424242; border-bottom:1px solid #E0E0E0;">{label}</td>
          <td style="padding:6px 12px; font-size:0.85rem; font-weight:700; color:#212121; border-bottom:1px solid #E0E0E0; text-align:right;">{value}{sub_html}</td>
        </tr>"""

    st.markdown(
        f"""
        <table style="width:100%; border-collapse:collapse; border:1px solid #E0E0E0; font-family:'Segoe UI',Arial,sans-serif;">
          {rows_html}
        </table>
        """,
        unsafe_allow_html=True,
    )
