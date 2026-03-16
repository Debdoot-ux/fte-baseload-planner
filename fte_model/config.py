from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

import pandas as pd


@dataclass
class NormStageData:
    """Peer benchmark data for one archetype in one TRL stage."""
    fte: float
    cost_myr: float
    duration_months: int
    technicians: float = 0.0
    direct_ratio: float = 0.0

    def _eff_fte(self, include_technicians: bool = True) -> float:
        return self.fte if include_technicians else max(self.fte - self.technicians, 0.0)

    def norm_value(self, metric: str, include_technicians: bool = True,
                   use_direct: bool = False) -> float:
        """Return the norm ratio for the chosen metric."""
        if use_direct:
            return self.direct_ratio
        if self.cost_myr <= 0:
            return 0.0
        eff = self._eff_fte(include_technicians)
        if metric == "py_per_myr":
            dur_years = self.duration_months / 12.0
            return (eff * dur_years) / self.cost_myr
        return eff / self.cost_myr


@dataclass
class NormSource:
    """A single peer company used as a staffing benchmark."""
    name: str
    data: Dict[Tuple[str, str], NormStageData] = field(default_factory=dict)
    input_mode: str = "detailed"  # "detailed" or "direct"


@dataclass
class NormsConfig:
    """Configuration for the staffing norms overlay."""
    sources: List[NormSource] = field(default_factory=list)
    selected_sources: List[str] = field(default_factory=lambda: ["Shell", "Chevron"])
    include_technicians: bool = True
    norm_metric: str = "fte_per_myr"  # "fte_per_myr" or "py_per_myr"

    @property
    def enabled(self) -> bool:
        return len(self.sources) > 0

    @property
    def metric_label(self) -> str:
        return "PY / MYR M" if self.norm_metric == "py_per_myr" else "FTE / MYR M"

    def _active_sources(self) -> List[NormSource]:
        if not self.selected_sources:
            return list(self.sources)
        name_set = set(self.selected_sources)
        return [s for s in self.sources if s.name in name_set]

    def combined_norms(self) -> Dict[Tuple[str, str], float]:
        """Compute the norm per (archetype, stage) bucket using the chosen metric."""
        active = self._active_sources()
        if not active:
            return {}

        inc_tech = self.include_technicians
        metric = self.norm_metric
        all_keys: set = set()
        for src in active:
            all_keys.update(src.data.keys())

        result: Dict[Tuple[str, str], float] = {}
        for key in all_keys:
            vals = [src.data[key].norm_value(metric, inc_tech,
                                             use_direct=(src.input_mode == "direct"))
                    for src in active if key in src.data]
            if not vals:
                continue
            result[key] = sum(vals) / len(vals)
        return result


@dataclass
class StageParams:
    """Parameters for one archetype in one pipeline stage."""

    duration_months: int
    cost_min: float
    cost_max: float
    fte_per_role: Dict[str, float] = field(default_factory=dict)

    @property
    def cost_millions(self) -> float:
        """Expected cost = midpoint of min/max range."""
        return (self.cost_min + self.cost_max) / 2.0


@dataclass
class Archetype:
    """A project archetype (e.g. Chemistry, Hardware, Algorithm)."""

    name: str
    portfolio_share: float
    stages: Dict[str, StageParams] = field(default_factory=dict)


@dataclass
class ModelConfig:
    """All user-configurable inputs for the FTE model."""

    total_budget_m: float = 400.0
    overhead_pct: float = 0.30
    start_year: int = 2026
    end_year: int = 2030
    horizon_end_year: int = 0  # 0 = disabled; when > end_year, annual summary extends to show wind-down

    budget_mode: str = "cashflow"  # "cashflow" or "commitment"

    pipeline_stages: List[str] = field(
        default_factory=lambda: ["TRL 1-4", "TRL 5-7"]
    )

    stage_mix: Dict[str, float] = field(
        default_factory=lambda: {"TRL 1-4": 0.20, "TRL 5-7": 0.80}
    )

    stage_conversion_rates: Dict[str, float] = field(
        default_factory=lambda: {"TRL 1-4": 0.40}
    )

    stage_mix_phase2: Dict[str, float] = field(default_factory=dict)
    phase2_start_year: int = 0

    intake_spread_months: int = 6
    utilization_rate: float = 1.0
    ramp_months: int = 0

    workforce_roles: List[str] = field(
        default_factory=lambda: ["Researcher", "Developer"]
    )
    contingency_pct: float = 0.0

    archetypes: List[Archetype] = field(default_factory=list)
    norms_config: NormsConfig = field(default_factory=NormsConfig)

    def __post_init__(self):
        if self.total_budget_m < 0:
            self.total_budget_m = 0.0
        self.overhead_pct = max(0.0, min(1.0, self.overhead_pct))
        self.intake_spread_months = max(1, min(12, self.intake_spread_months))
        if self.end_year < self.start_year:
            self.end_year = self.start_year
        if self.archetypes:
            share_sum = sum(a.portfolio_share for a in self.archetypes)
            if abs(share_sum - 1.0) > 0.01:
                warnings.warn(
                    f"Archetype portfolio shares sum to {share_sum:.2f}, expected 1.0",
                    stacklevel=2,
                )
        if self.stage_mix:
            mix_sum = sum(self.stage_mix.values())
            if abs(mix_sum - 1.0) > 0.01:
                warnings.warn(
                    f"Stage mix values sum to {mix_sum:.2f}, expected 1.0",
                    stacklevel=2,
                )

    @property
    def all_roles(self) -> List[str]:
        """Unique role names across all archetypes, preserving first-seen order."""
        seen: set = set()
        roles: List[str] = []
        for arch in self.archetypes:
            for sp in arch.stages.values():
                for role in sp.fte_per_role:
                    if role not in seen:
                        seen.add(role)
                        roles.append(role)
        return roles if roles else list(self.workforce_roles)


@dataclass
class ModelResult:
    """Output container returned by the calculation engine."""

    monthly: pd.DataFrame
    annual_summary: pd.DataFrame
    steady_state_avg: float
    steady_state_min_month: float
    steady_state_max_month: float
    projects_per_year: float
    yearly_projects: Dict[int, float] = field(default_factory=dict)
    # Cost sensitivity band (populated when cost_min != cost_max)
    cost_low_ss_avg: float = 0.0
    cost_high_ss_avg: float = 0.0
    cost_low_annual: pd.DataFrame = field(default_factory=pd.DataFrame)
    cost_high_annual: pd.DataFrame = field(default_factory=pd.DataFrame)
    # Norms overlay (populated when norms_config is enabled)
    norms_annual: pd.DataFrame = field(default_factory=pd.DataFrame)
    norms_breakdown: pd.DataFrame = field(default_factory=pd.DataFrame)
    combined_norms: Dict = field(default_factory=dict)
