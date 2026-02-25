from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List

import pandas as pd


@dataclass
class StageParams:
    """Parameters for one archetype in one pipeline stage."""

    duration_months: int
    cost_millions: float
    fte_research: float
    fte_developer: float


@dataclass
class Archetype:
    """A project archetype (e.g. Chemistry, Hardware, Software)."""

    name: str
    portfolio_share: float
    stages: Dict[str, StageParams] = field(default_factory=dict)


@dataclass
class ModelConfig:
    """All user-configurable inputs for the FTE model."""

    total_budget_m: float = 400.0
    overhead_pct: float = 0.30
    start_year: int = 2026
    end_year: int = 2029

    pipeline_stages: List[str] = field(
        default_factory=lambda: ["TRL 1-4", "TRL 5-7"]
    )

    stage_mix: Dict[str, float] = field(
        default_factory=lambda: {"TRL 1-4": 0.20, "TRL 5-7": 0.80}
    )

    stage_conversion_rates: Dict[str, float] = field(
        default_factory=lambda: {"TRL 1-4": 0.50}
    )

    intake_spread_months: int = 6
    utilization_rate: float = 1.0
    ramp_months: int = 0

    archetypes: List[Archetype] = field(default_factory=list)


@dataclass
class ModelResult:
    """Output container returned by the calculation engine."""

    monthly: pd.DataFrame
    annual_summary: pd.DataFrame
    steady_state_avg: float
    steady_state_min_month: float
    steady_state_max_month: float
    projects_per_year: float
