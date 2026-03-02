"""Baseline assumptions for the FTE model."""

from config import Archetype, ModelConfig, StageParams


def default_baseline() -> ModelConfig:
    chemistry = Archetype(
        name="Chemistry",
        portfolio_share=0.20,
        stages={
            "TRL 1-4": StageParams(
                duration_months=7,
                cost_millions=4.0,
                fte_research=3.5,
                fte_developer=1.5,
            ),
            "TRL 5-7": StageParams(
                duration_months=21,
                cost_millions=8.0,
                fte_research=1.5,
                fte_developer=3.5,
            ),
        },
    )

    hardware_mechanical = Archetype(
        name="Hardware: Mechanical",
        portfolio_share=0.30,
        stages={
            "TRL 1-4": StageParams(
                duration_months=9,
                cost_millions=5.0,
                fte_research=3.5,
                fte_developer=1.5,
            ),
            "TRL 5-7": StageParams(
                duration_months=21,
                cost_millions=35.0,
                fte_research=1.5,
                fte_developer=3.5,
            ),
        },
    )

    hardware_process = Archetype(
        name="Hardware: Process",
        portfolio_share=0.30,
        stages={
            "TRL 1-4": StageParams(
                duration_months=24,
                cost_millions=15.0,
                fte_research=6.5,
                fte_developer=1.5,
            ),
            "TRL 5-7": StageParams(
                duration_months=75,
                cost_millions=50.0,
                fte_research=1.5,
                fte_developer=6.5,
            ),
        },
    )

    software = Archetype(
        name="Algorithm (Software)",
        portfolio_share=0.20,
        stages={
            "TRL 1-4": StageParams(
                duration_months=9,
                cost_millions=2.0,
                fte_research=0.5,
                fte_developer=0.5,
            ),
            "TRL 5-7": StageParams(
                duration_months=9,
                cost_millions=5.0,
                fte_research=0.5,
                fte_developer=0.5,
            ),
        },
    )

    return ModelConfig(
        total_budget_m=400.0,
        overhead_pct=0.30,
        start_year=2026,
        end_year=2030,
        budget_mode="cashflow",
        pipeline_stages=["TRL 1-4", "TRL 5-7"],
        stage_mix={"TRL 1-4": 0.20, "TRL 5-7": 0.80},
        stage_conversion_rates={"TRL 1-4": 0.40},
        stage_mix_phase2={"TRL 1-4": 0.40, "TRL 5-7": 0.60},
        phase2_start_year=2028,
        intake_spread_months=6,
        utilization_rate=1.0,
        ramp_months=0,
        archetypes=[chemistry, hardware_mechanical, hardware_process, software],
    )
