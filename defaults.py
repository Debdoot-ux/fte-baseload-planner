"""Baseline assumptions (PETRONAS R&D context)."""

from config import Archetype, ModelConfig, StageParams


def petronas_baseline() -> ModelConfig:
    chemistry = Archetype(
        name="Chemistry",
        portfolio_share=0.15,
        stages={
            "TRL 1-4": StageParams(
                duration_months=7,
                cost_millions=6.5,
                fte_research=3.5,
                fte_developer=1.5,
            ),
            "TRL 5-7": StageParams(
                duration_months=12,
                cost_millions=12.5,
                fte_research=1.5,
                fte_developer=3.5,
            ),
        },
    )

    hardware = Archetype(
        name="Process (Hardware)",
        portfolio_share=0.70,
        stages={
            "TRL 1-4": StageParams(
                duration_months=9,
                cost_millions=12.5,
                fte_research=6.5,
                fte_developer=1.5,
            ),
            "TRL 5-7": StageParams(
                duration_months=15,
                cost_millions=15.0,
                fte_research=1.5,
                fte_developer=6.5,
            ),
        },
    )

    software = Archetype(
        name="Algorithm (Software)",
        portfolio_share=0.15,
        stages={
            "TRL 1-4": StageParams(
                duration_months=6,
                cost_millions=4.25,
                fte_research=0.5,
                fte_developer=0.5,
            ),
            "TRL 5-7": StageParams(
                duration_months=6,
                cost_millions=4.25,
                fte_research=0.5,
                fte_developer=0.5,
            ),
        },
    )

    return ModelConfig(
        total_budget_m=400.0,
        overhead_pct=0.30,
        start_year=2026,
        end_year=2029,
        pipeline_stages=["TRL 1-4", "TRL 5-7"],
        stage_mix={"TRL 1-4": 0.20, "TRL 5-7": 0.80},
        stage_conversion_rates={"TRL 1-4": 0.50},
        intake_spread_months=6,
        utilization_rate=1.0,
        ramp_months=0,
        archetypes=[chemistry, hardware, software],
    )
