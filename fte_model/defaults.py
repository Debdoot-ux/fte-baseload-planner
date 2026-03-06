"""Baseline assumptions for the FTE model."""

from config import Archetype, ModelConfig, StageParams


def default_baseline() -> ModelConfig:
    chemistry = Archetype(
        name="Chemistry",
        portfolio_share=0.20,
        stages={
            "TRL 1-4": StageParams(
                duration_months=7,
                cost_min=4.0,
                cost_max=4.0,
                fte_per_role={"Researcher": 3.5, "Developer": 1.5},
            ),
            "TRL 5-7": StageParams(
                duration_months=21,
                cost_min=8.0,
                cost_max=8.0,
                fte_per_role={"Researcher": 1.5, "Developer": 3.5},
            ),
        },
    )

    hardware_mechanical = Archetype(
        name="Hardware: Mechanical",
        portfolio_share=0.30,
        stages={
            "TRL 1-4": StageParams(
                duration_months=9,
                cost_min=5.0,
                cost_max=5.0,
                fte_per_role={"Researcher": 3.5, "Developer": 1.5},
            ),
            "TRL 5-7": StageParams(
                duration_months=21,
                cost_min=35.0,
                cost_max=35.0,
                fte_per_role={"Researcher": 1.5, "Developer": 3.5},
            ),
        },
    )

    hardware_process = Archetype(
        name="Hardware: Process",
        portfolio_share=0.30,
        stages={
            "TRL 1-4": StageParams(
                duration_months=24,
                cost_min=15.0,
                cost_max=15.0,
                fte_per_role={"Researcher": 6.5, "Developer": 1.5},
            ),
            "TRL 5-7": StageParams(
                duration_months=75,
                cost_min=50.0,
                cost_max=50.0,
                fte_per_role={"Researcher": 1.5, "Developer": 6.5},
            ),
        },
    )

    algorithm = Archetype(
        name="Algorithm",
        portfolio_share=0.20,
        stages={
            "TRL 1-4": StageParams(
                duration_months=9,
                cost_min=2.0,
                cost_max=2.0,
                fte_per_role={"Researcher": 0.5, "Developer": 0.5},
            ),
            "TRL 5-7": StageParams(
                duration_months=9,
                cost_min=5.0,
                cost_max=5.0,
                fte_per_role={"Researcher": 0.5, "Developer": 0.5},
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
        workforce_roles=["Researcher", "Developer"],
        archetypes=[chemistry, hardware_mechanical, hardware_process, algorithm],
    )
