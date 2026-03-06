from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List

import pandas as pd


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
