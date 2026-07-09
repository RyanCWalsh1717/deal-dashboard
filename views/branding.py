"""GRP-branded hero banner — same green-gradient look as ga-automation's app.py,
so both internal apps read as one family. Renders via st.components.v1.html()
(a sandboxed iframe) rather than st.markdown/st.html, which have both been
unreliable for complex HTML in recent Streamlit versions.
"""

from __future__ import annotations

import base64
from pathlib import Path
from typing import List, Optional

ASSETS_DIR = Path(__file__).parent.parent / "assets"

_MIME_TYPES = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "svg": "image/svg+xml", "webp": "image/webp"}


def _img_b64(fname: str) -> Optional[str]:
    path = ASSETS_DIR / fname
    if not path.exists():
        return None
    raw = base64.b64encode(path.read_bytes()).decode()
    ext = fname.rsplit(".", 1)[-1].lower()
    return f"data:{_MIME_TYPES.get(ext, 'image/png')};base64,{raw}"


_LOGO_SRC = _img_b64("grp_logo.png") or _img_b64("grp_logo.svg")


def _hero_photo_src(property_code: str) -> Optional[str]:
    if not property_code:
        return None
    for ext in (".jpg", ".jpeg", ".png", ".webp"):
        src = _img_b64(f"{property_code}_hero{ext}")
        if src:
            return src
    return None


def render_hero(title: str, subtitle: str = "", badges: Optional[List[str]] = None, photo_code: str = "") -> None:
    import streamlit.components.v1 as stc

    badges = badges or []
    photo_src = _hero_photo_src(photo_code)
    photo_html = f'<img src="{photo_src}" class="grp-hero-photo" alt="{title}"/>' if photo_src else ""
    logo_html = (
        f'<img src="{_LOGO_SRC}" style="max-width:140px;max-height:60px;" alt="GRP Logo"/>'
        if _LOGO_SRC
        else '<div class="grp-logo-text">Greatland<br>Realty<br>Partners</div>'
    )
    badge_html = " ".join(f'<span class="grp-badge">{b}</span>' for b in badges)

    stc.html(
        f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  html, body {{
    margin: 0; padding: 0; overflow: hidden;
    font-family: 'Segoe UI', Arial, sans-serif;
    background: transparent;
  }}
  .grp-hero {{
    background: linear-gradient(135deg, #1A5C22 0%, #2E7D32 100%);
    border-radius: 10px;
    padding: 0;
    overflow: hidden;
    box-shadow: 0 4px 12px rgba(0,0,0,0.18);
    display: flex;
    align-items: stretch;
    min-height: 130px;
    margin: 0;
  }}
  .grp-hero-photo {{
    width: 240px; min-width: 240px;
    object-fit: cover;
    border-radius: 10px 0 0 10px;
    display: block;
  }}
  .grp-hero-body {{
    padding: 18px 24px;
    flex: 1;
    display: flex;
    flex-direction: column;
    justify-content: center;
  }}
  .grp-hero-title {{
    color: #ffffff;
    font-size: 1.45rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    margin: 0 0 4px 0;
    line-height: 1.2;
  }}
  .grp-hero-sub {{
    color: #C8E6C9;
    font-size: 0.85rem;
    margin: 0 0 8px 0;
    font-weight: 400;
  }}
  .grp-hero-badges {{
    display: flex;
    gap: 7px;
    flex-wrap: wrap;
    margin-top: 3px;
  }}
  .grp-badge {{
    background: rgba(255,255,255,0.18);
    border: 1px solid rgba(255,255,255,0.35);
    border-radius: 20px;
    padding: 2px 11px;
    color: #ffffff;
    font-size: 0.73rem;
    font-weight: 500;
    white-space: nowrap;
  }}
  .grp-hero-logo {{
    padding: 18px 20px;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    min-width: 150px;
  }}
  .grp-logo-text {{
    color: rgba(255,255,255,0.85);
    font-size: 0.68rem;
    text-align: right;
    line-height: 1.4;
    font-weight: 500;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }}
</style>
</head>
<body>
<div class="grp-hero">
  {photo_html}
  <div class="grp-hero-body">
    <div class="grp-hero-title">{title}</div>
    <div class="grp-hero-sub">{subtitle}</div>
    <div class="grp-hero-badges">{badge_html}</div>
  </div>
  <div class="grp-hero-logo">{logo_html}</div>
</div>
</body>
</html>
""",
        height=150,
        scrolling=False,
    )
