"""Per-property configuration, loaded from `data/<property_code>/config.yaml`.

Extends the YAML-per-property + auto-discovery pattern from ga-automation's
`property_config.py` with JV/ownership-tier fields. Not cross-imported from that
project — separate repo, own copy of the pattern.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import yaml

TEMPLATE_CODE = "TEMPLATE"


@dataclass
class InvestorShare:
    display_name: str
    legal_entity: str = ""
    ownership_pct: float = 0.0
    role: str = ""  # 'LP' | 'co-GP' | 'co-GP member' | 'Sponsor'
    is_greatland_affiliate: bool = False
    sub_tier: Optional[str] = None  # tier_id of a nested tier, or None (terminal investor)


@dataclass
class OwnershipTier:
    tier_id: str
    distributing_entity: str = ""
    parent_tier: Optional[str] = None
    display_label: str = ""
    investors: List[InvestorShare] = field(default_factory=list)

    def label(self) -> str:
        return self.display_label or self.tier_id


@dataclass
class LoanRef:
    tranche_name: str
    lender: str = ""
    abstract_file: str = ""


@dataclass
class JVDocumentRef:
    name: str
    abstract_file: str = ""


@dataclass
class PropertyConfig:
    property_code: str
    property_name: str = ""
    property_display_name: str = ""
    property_address: str = ""
    property_type: str = ""
    market: str = ""
    submarket: str = ""
    state: str = ""
    management_company: str = "Greatland Realty Partners"
    active: bool = True

    source_workbook_path_hint: str = "source_files/distribution_workbook.xlsx"
    sheet_map: Dict[str, str] = field(default_factory=dict)

    # The dashboard's own `property_code` (used for the data/<code>/ folder name) is
    # just an internal identifier — it does NOT necessarily match the property code(s)
    # used in Yardi-sourced exports (trial balances, rent rolls). Confirmed for
    # Revolution Labs: the dashboard folder is "revlabspm" but the real Yardi exports
    # say "revlabpm" (property entity), "revlabvn" (venture/JV entity), "bh1050jv" (co-GP
    # entity) — none of which match "revlabspm" as a substring. Anything that needs to
    # filter a multi-entity source file (e.g. a trial balance covering several
    # properties) down to just this property should match against yardi_codes, not
    # property_code.
    yardi_codes: List[str] = field(default_factory=list)

    ownership_tiers: List[OwnershipTier] = field(default_factory=list)
    loans: List[LoanRef] = field(default_factory=list)
    jv_documents: List[JVDocumentRef] = field(default_factory=list)

    def display(self) -> str:
        return self.property_display_name or self.property_name or self.property_code

    def top_level_tier(self) -> Optional[OwnershipTier]:
        for tier in self.ownership_tiers:
            if tier.parent_tier is None:
                return tier
        return None

    def get_tier(self, tier_id: str) -> Optional[OwnershipTier]:
        for tier in self.ownership_tiers:
            if tier.tier_id == tier_id:
                return tier
        return None

    def investor_display_names(self) -> List[str]:
        return [inv.display_name for tier in self.ownership_tiers for inv in tier.investors]

    def workbook_path(self, data_dir: str = "data") -> Path:
        return Path(data_dir) / self.property_code / self.source_workbook_path_hint

    def abstract_path(self, abstract_file: str, data_dir: str = "data") -> Path:
        return Path(data_dir) / self.property_code / abstract_file

    @classmethod
    def load(cls, property_code: str, data_dir: str = "data") -> "PropertyConfig":
        config_path = Path(data_dir) / property_code / "config.yaml"
        raw = yaml.safe_load(config_path.read_text()) or {}
        return cls._from_dict(property_code, raw)

    @classmethod
    def _from_dict(cls, property_code: str, raw: dict) -> "PropertyConfig":
        tiers = [
            OwnershipTier(
                tier_id=t["tier_id"],
                distributing_entity=t.get("distributing_entity", ""),
                parent_tier=t.get("parent_tier"),
                display_label=t.get("display_label", ""),
                investors=[InvestorShare(**inv) for inv in t.get("investors", [])],
            )
            for t in raw.get("ownership_tiers", [])
        ]
        loans = [LoanRef(**loan) for loan in raw.get("loans", [])]
        jv_documents = [JVDocumentRef(**doc) for doc in raw.get("jv_documents", [])]

        known_fields = {
            "property_name",
            "property_display_name",
            "property_address",
            "property_type",
            "market",
            "submarket",
            "state",
            "management_company",
            "active",
            "source_workbook_path_hint",
            "sheet_map",
            "yardi_codes",
        }
        scalar_kwargs = {k: v for k, v in raw.items() if k in known_fields}

        return cls(
            property_code=property_code,
            ownership_tiers=tiers,
            loans=loans,
            jv_documents=jv_documents,
            **scalar_kwargs,
        )


def discover_properties(data_dir: str = "data") -> List[PropertyConfig]:
    """Auto-discovers every `data/<code>/config.yaml` except the TEMPLATE scaffold."""
    root = Path(data_dir)
    if not root.exists():
        return []
    configs = []
    for child in sorted(root.iterdir()):
        if not child.is_dir() or child.name == TEMPLATE_CODE:
            continue
        if (child / "config.yaml").exists():
            configs.append(PropertyConfig.load(child.name, data_dir=data_dir))
    return configs


def list_active_properties(data_dir: str = "data") -> List[PropertyConfig]:
    return [cfg for cfg in discover_properties(data_dir) if cfg.active]
