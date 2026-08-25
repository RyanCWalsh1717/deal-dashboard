"""Save a PropertyConfig to YAML (local disk + optional GitHub API), so a new
property can be onboarded from the Properties tab without anyone hand-editing
YAML or touching GitHub directly.

Ported from ga-automation's own `property_writer.py` (same save-priority
design, same secrets shape) — not cross-imported, this is Deal Dashboard's own
copy, adapted to this app's PropertyConfig schema (ownership tiers/loans/JV
documents instead of accrual accounts/management fees/bank accounts).

Save priority:
  1. GitHub API — if token + repo in st.secrets or environment vars. Triggers
     an automatic Streamlit Cloud redeploy, so this is what makes a
     live-hosted app's own filesystem writes (below) actually stick —
     Streamlit Cloud's filesystem is ephemeral and resets on redeploy.
  2. Local disk — always attempted; this is what makes it work at all for
     local/dev use, and is a harmless no-op cache on top of GitHub for hosted use.
  3. Download — the YAML text is always returned too, so the UI can offer a
     manual-download fallback regardless of whether either save above succeeds.

Secrets expected in .streamlit/secrets.toml or Streamlit Cloud settings:
    [github]
    token = "ghp_..."
    repo  = "RyanCWalsh1717/deal-dashboard"   # owner/repo
"""

from __future__ import annotations

import os
from typing import Optional, Tuple

import yaml


def config_to_yaml(data: dict) -> str:
    """Render a property config dict as a YAML string. Block style, key order
    preserved (sort_keys=False) so it reads the same shape as the TEMPLATE."""
    return yaml.dump(data, default_flow_style=False, allow_unicode=True, sort_keys=False, width=80)


def build_config_dict(
    property_code: str,
    property_name: str = "",
    property_display_name: str = "",
    property_address: str = "",
    property_type: str = "",
    market: str = "",
    submarket: str = "",
    state: str = "",
    management_company: str = "Greatland Realty Partners",
    active: bool = True,
    source_workbook_path_hint: str = "source_files/distribution_workbook.xlsx",
    sheet_map: Optional[dict] = None,
    yardi_codes: Optional[list] = None,
    ownership_tiers: Optional[list] = None,
    loans: Optional[list] = None,
    jv_documents: Optional[list] = None,
) -> dict:
    """Builds the dict PropertyConfig._from_dict() expects (same field names,
    same shape as data/TEMPLATE/config.yaml) — property_code is NOT included
    in the dict itself (it comes from the parent directory name, same as
    every other config.yaml in this app)."""
    return {
        "property_name": property_name,
        "property_display_name": property_display_name,
        "property_address": property_address,
        "property_type": property_type,
        "market": market,
        "submarket": submarket,
        "state": state,
        "management_company": management_company,
        "active": active,
        "source_workbook_path_hint": source_workbook_path_hint,
        "sheet_map": sheet_map or {},
        "yardi_codes": yardi_codes or [],
        "ownership_tiers": ownership_tiers or [],
        "loans": loans or [],
        "jv_documents": jv_documents or [],
    }


def save_local(property_code: str, yaml_content: str, data_dir: str) -> Tuple[bool, str]:
    """Writes config.yaml to data/{property_code}/ on local disk."""
    try:
        folder = os.path.join(data_dir, property_code)
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "config.yaml")
        with open(path, "w", encoding="utf-8") as f:
            f.write(yaml_content)
        return True, path
    except Exception as e:
        return False, str(e)


_HERO_EXTS = (".jpg", ".jpeg", ".png", ".webp")


def save_hero_photo_local(property_code: str, image_bytes: bytes, ext: str, assets_dir: str) -> Tuple[bool, str]:
    """Saves a property's hero photo to assets/{property_code}_hero{ext} — the
    exact filename views/branding.py's _hero_photo_src() looks for (unlike
    ga-automation's own save_image_local(), which writes under data/<code>/;
    Deal Dashboard's hero photos have always lived in a flat assets/ folder,
    see e.g. assets/revlabspm_hero.jpg). Deletes any OTHER extension's hero
    file for this property first, so an old photo can't keep winning
    _hero_photo_src()'s fixed jpg→jpeg→png→webp search order after it's
    been replaced with a different format."""
    try:
        folder = os.path.join(assets_dir)
        os.makedirs(folder, exist_ok=True)
        for other_ext in _HERO_EXTS:
            if other_ext != ext:
                stale = os.path.join(folder, f"{property_code}_hero{other_ext}")
                if os.path.exists(stale):
                    os.remove(stale)
        path = os.path.join(folder, f"{property_code}_hero{ext}")
        with open(path, "wb") as f:
            f.write(image_bytes)
        return True, path
    except Exception as e:
        return False, str(e)


def save_hero_photo_to_github(property_code: str, image_bytes: bytes, ext: str) -> Tuple[bool, str]:
    """Uploads a property's hero photo to assets/{property_code}_hero{ext} in
    the GitHub repo. Doesn't delete other-extension files in GitHub (the
    contents API only writes/updates one path at a time) — a stale photo left
    behind there after a format change is a minor, self-resolving cosmetic
    gap, not a correctness issue, since the local save above already prevents
    it for local runs."""
    import base64

    try:
        import requests
    except ImportError:
        return False, "requests library not available"

    token, repo = _github_credentials()
    if not token or not repo:
        return False, "GitHub token/repo not configured in secrets"

    path = f"assets/{property_code}_hero{ext}"
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

    sha = None
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha = r.json().get("sha")
    except Exception:
        pass

    payload: dict = {
        "message": f"Upload hero photo: {property_code}_hero{ext}",
        "content": base64.b64encode(image_bytes).decode("ascii"),
    }
    if sha:
        payload["sha"] = sha

    try:
        r = requests.put(url, json=payload, headers=headers, timeout=30)
        if r.status_code in (200, 201):
            return True, "Photo saved to GitHub. Hero banner updates after ~2 min redeploy."
        return False, f"GitHub API returned {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, str(e)


def _github_credentials() -> Tuple[str, str]:
    """Returns (token, repo) from Streamlit secrets or environment variables,
    or ('', '') if neither is configured."""
    token, repo = "", ""
    try:
        import streamlit as st

        token = st.secrets.get("github", {}).get("token", "")
        repo = st.secrets.get("github", {}).get("repo", "")
    except Exception:
        pass
    token = token or os.environ.get("GITHUB_TOKEN", "")
    repo = repo or os.environ.get("GITHUB_REPO", "")
    return token, repo


def github_configured() -> bool:
    token, repo = _github_credentials()
    return bool(token and repo)


def save_to_github(property_code: str, yaml_content: str, commit_message: str = "") -> Tuple[bool, str]:
    """Writes config.yaml to the GitHub repo via the REST API — creates the
    file if it doesn't exist yet (a brand-new property), updates it if it does
    (editing an existing one). Reads credentials from
    st.secrets['github']['token']/['repo'] or GITHUB_TOKEN/GITHUB_REPO."""
    import base64

    try:
        import requests
    except ImportError:
        return False, "requests library not available"

    token, repo = _github_credentials()
    if not token or not repo:
        return False, "GitHub token/repo not configured in secrets"

    path = f"data/{property_code}/config.yaml"
    url = f"https://api.github.com/repos/{repo}/contents/{path}"
    headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

    sha = None
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha = r.json().get("sha")
    except Exception:
        pass

    payload: dict = {
        "message": commit_message or f"Add/update property config: {property_code}",
        "content": base64.b64encode(yaml_content.encode("utf-8")).decode("ascii"),
    }
    if sha:
        payload["sha"] = sha

    try:
        r = requests.put(url, json=payload, headers=headers, timeout=15)
        if r.status_code in (200, 201):
            action = "updated" if sha else "created"
            return True, f"Config {action} in GitHub. Streamlit Cloud will redeploy in ~2 minutes."
        return False, f"GitHub API returned {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, str(e)
