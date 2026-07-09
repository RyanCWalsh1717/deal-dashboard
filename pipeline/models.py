"""Shared dataclasses for parsed distribution-workbook data.

Every parser in `pipeline/parsers/` returns instances of these types rather than
raw dicts, so `views/` code and `app.py` never need to know about workbook layout.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Dict, List, Optional


@dataclass
class CashFlowLine:
    account_code: str
    account_label: str
    monthly_values: Dict[date, float] = field(default_factory=dict)


@dataclass
class CashFlowResult:
    property_code: str
    period_columns: List[date] = field(default_factory=list)
    lines: List[CashFlowLine] = field(default_factory=list)
    subtotals: Dict[str, Dict[date, float]] = field(default_factory=dict)
    # "actual" | "budget" | "reforecast" per period column, when the tab labels
    # it (confirmed in the Revolution Labs workbook: 2026 is Jan-May Actuals,
    # Jun-Dec Budget — a single year can blend both scenarios in one row).
    period_scenario: Dict[date, str] = field(default_factory=dict)

    def subtotal(self, label: str, period: date) -> Optional[float]:
        return self.subtotals.get(label, {}).get(period)

    def as_of(self, period: date) -> Dict[str, float]:
        return {
            line.account_code: line.monthly_values[period]
            for line in self.lines
            if period in line.monthly_values
        }

    def latest_period(self) -> Optional[date]:
        return max(self.period_columns) if self.period_columns else None

    def last_actual_period(self) -> Optional[date]:
        """The most recent period genuinely labeled "actual" — NOT just the
        latest column with data, since forecast/budget columns for the
        current year live in the same row and would otherwise be mistaken
        for real actuals."""
        actual_periods = [p for p, s in self.period_scenario.items() if s.startswith("actual")]
        if actual_periods:
            return max(actual_periods)
        return self.latest_period() if not self.period_scenario else None

    def total_cash_position(self, period: date, cash_prefixes=("11",)) -> float:
        return sum(
            line.monthly_values.get(period, 0.0)
            for line in self.lines
            if line.account_code[: len(cash_prefixes[0])] in cash_prefixes
        )


@dataclass
class EquityPosition:
    entity_name: str
    entity_code: str = ""
    as_of_period: str = ""
    balance_sheet: Dict[str, float] = field(default_factory=dict)
    contributions_by_partner: Dict[str, float] = field(default_factory=dict)
    distributions_by_partner: Dict[str, float] = field(default_factory=dict)

    def _total(self, *needles: str) -> Optional[float]:
        for key, value in self.balance_sheet.items():
            if all(n in key for n in needles):
                return value
        return None

    @property
    def total_cash(self) -> Optional[float]:
        return self._total("TOTAL", "CASH")

    @property
    def total_contributions(self) -> Optional[float]:
        return self.balance_sheet.get("TOTAL CONTRIBUTIONS")

    @property
    def total_distributions(self) -> Optional[float]:
        return self.balance_sheet.get("TOTAL DISTRIBUTIONS")


@dataclass
class DebtTranche:
    tranche_name: str
    outstanding_balance: float
    interest_rate: float
    as_of: Optional[date] = None
    paydown: float = 0.0
    interest_carry: float = 0.0


@dataclass
class DebtSummary:
    property_code: str
    tranches: List[DebtTranche] = field(default_factory=list)
    as_of: Optional[date] = None
    sofr_as_of: Optional[float] = None

    @property
    def total_outstanding(self) -> float:
        return sum(t.outstanding_balance for t in self.tranches)


@dataclass
class InvestorDistribution:
    display_name: str
    legal_entity: str
    ownership_pct: float
    distribution_amount: float
    distribution_pct_of_contributions: Optional[float] = None
    contributions_to_date: Optional[float] = None
    distributions_to_date: Optional[float] = None
    net_capital_after: Optional[float] = None


@dataclass
class WaterfallTier:
    tier_id: str
    distributing_entity: str
    as_of_label: str = ""
    cash_projected: Optional[float] = None
    cash_holdbacks: Dict[str, float] = field(default_factory=dict)
    net_cash_available: Optional[float] = None
    distribution_recommendation: Optional[float] = None
    investors: List[InvestorDistribution] = field(default_factory=list)


@dataclass
class DistributionWaterfall:
    property_code: str
    tiers: Dict[str, WaterfallTier] = field(default_factory=dict)


@dataclass
class ProjectedWaterfallTier:
    """Placeholder shape for the forward-looking (promote/hurdle) waterfall —
    no parser populates this until the leasing/investment model is finalized."""

    tier_id: str
    hurdle_tiers: List[dict] = field(default_factory=list)


@dataclass
class ProjectedDistributionWaterfall:
    property_code: str
    tiers: Dict[str, ProjectedWaterfallTier] = field(default_factory=dict)


@dataclass
class BudgetLine:
    account_code: str
    account_label: str
    budget_value: Optional[float]
    actual_value: Optional[float]

    @property
    def variance_dollar(self) -> Optional[float]:
        if self.budget_value is None or self.actual_value is None:
            return None
        return self.actual_value - self.budget_value

    @property
    def variance_pct(self) -> Optional[float]:
        if self.budget_value in (None, 0) or self.actual_value is None:
            return None
        return (self.actual_value - self.budget_value) / abs(self.budget_value)

    @property
    def missing_side(self) -> Optional[str]:
        if self.budget_value is None:
            return "budget"
        if self.actual_value is None:
            return "actual"
        return None


@dataclass
class BudgetComparisonResult:
    property_code: str
    period: Optional[date]
    lines: List[BudgetLine] = field(default_factory=list)

    def total_variance(self) -> float:
        return sum(
            line.variance_dollar for line in self.lines if line.variance_dollar is not None
        )


@dataclass
class DistributionWorkbookResult:
    property_code: str
    source_path: str
    parsed_at: datetime
    cash_flow: Optional[CashFlowResult] = None
    equity: Dict[str, EquityPosition] = field(default_factory=dict)
    debt: Optional[DebtSummary] = None
    waterfall: Optional[DistributionWaterfall] = None
    projected_waterfall: Optional[ProjectedDistributionWaterfall] = None
    budget_comparison: Optional[BudgetComparisonResult] = None
    budget_summary: Optional[BudgetComparisonResult] = None


@dataclass
class CashAccountBalance:
    """One account box for the Cash tab — e.g. Operating Cash, DACA, an escrow.
    `source` records where it came from (equity_tab / trial_balance / loan_statement /
    placeholder) since these boxes may be assembled from several different files."""

    label: str
    balance: Optional[float]
    account_code: str = ""
    source: str = ""
    as_of: Optional[date] = None


@dataclass
class RentRollLine:
    building: str = ""
    floor: str = ""
    unit_code: str = ""
    unit_area: Optional[float] = None
    tenant_name: str = ""
    lease_from: Optional[date] = None
    lease_to: Optional[date] = None
    term_months: Optional[float] = None
    lease_area: Optional[float] = None
    annual_rent: Optional[float] = None
    annual_rent_psf: Optional[float] = None
    lease_type: str = ""
    is_vacant: bool = False


@dataclass
class RentRollResult:
    property_code: str
    as_of: Optional[date] = None
    lines: List[RentRollLine] = field(default_factory=list)

    @property
    def total_leased_sf(self) -> float:
        return sum(l.unit_area or 0.0 for l in self.lines if not l.is_vacant)

    @property
    def total_vacant_sf(self) -> float:
        return sum(l.unit_area or 0.0 for l in self.lines if l.is_vacant)

    @property
    def total_annual_rent(self) -> float:
        return sum(l.annual_rent or 0.0 for l in self.lines if not l.is_vacant)

    @property
    def occupancy_pct(self) -> Optional[float]:
        total = self.total_leased_sf + self.total_vacant_sf
        return (self.total_leased_sf / total) if total else None


@dataclass
class LoanStatement:
    """One month's Berkadia-format loan servicer statement for one tranche."""

    tranche_name: str
    loan_number: str = ""
    interest_rate: Optional[float] = None  # decimal, e.g. 0.0675248 — the real all-in rate
    as_of: Optional[date] = None
    principal_balance: Optional[float] = None
    interest_paid_ytd: Optional[float] = None
    tax_escrow_balance: Optional[float] = None
    insurance_escrow_balance: Optional[float] = None
    reserve_balance: Optional[float] = None
    total_payment_due: Optional[float] = None


@dataclass
class PortfolioSummaryRow:
    property_code: str
    display_name: str
    address: str = ""
    market: str = ""
    investor_names: List[str] = field(default_factory=list)
    total_cash: Optional[float] = None
    total_debt_outstanding: Optional[float] = None
    last_distribution_amount: Optional[float] = None
    last_distribution_as_of: Optional[str] = None
