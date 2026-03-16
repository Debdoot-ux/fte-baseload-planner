"""Baseline assumptions for the FTE model."""

from config import Archetype, ModelConfig, NormSource, NormStageData, NormsConfig, StageParams


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
        utilization_rate=0.70,
        ramp_months=0,
        workforce_roles=["Researcher", "Developer"],
        archetypes=[chemistry, hardware_mechanical, hardware_process, algorithm],
        norms_config=default_norms_config(),
    )


def default_norm_sources() -> list[NormSource]:
    """Shell, Chevron, and BASF benchmark data from v5 normalization workbook."""
    N = NormStageData
    shell = NormSource(name="Shell", data={
        ("Chemistry", "TRL 1-4"):            N(fte=6.5,  cost_myr=40.0,  duration_months=48),
        ("Chemistry", "TRL 5-7"):            N(fte=10.0, cost_myr=200.0, duration_months=72),
        ("Hardware: Mechanical", "TRL 1-4"):  N(fte=12.0, cost_myr=40.0,  duration_months=48),
        ("Hardware: Mechanical", "TRL 5-7"):  N(fte=17.5, cost_myr=100.0, duration_months=72),
        ("Hardware: Process", "TRL 1-4"):     N(fte=6.5,  cost_myr=80.0,  duration_months=60),
        ("Hardware: Process", "TRL 5-7"):     N(fte=15.0, cost_myr=200.0, duration_months=78),
        ("Algorithm", "TRL 1-4"):            N(fte=4.5,  cost_myr=8.0,   duration_months=18),
        ("Algorithm", "TRL 5-7"):            N(fte=10.0, cost_myr=24.0,  duration_months=18),
    })
    chevron = NormSource(name="Chevron", data={
        ("Chemistry", "TRL 1-4"):            N(fte=2.5,  cost_myr=6.0,   duration_months=48),
        ("Chemistry", "TRL 5-7"):            N(fte=15.5, cost_myr=140.0, duration_months=72),
        ("Hardware: Mechanical", "TRL 1-4"):  N(fte=8.0,  cost_myr=6.0,   duration_months=48),
        ("Hardware: Mechanical", "TRL 5-7"):  N(fte=27.5, cost_myr=140.0, duration_months=72),
        ("Hardware: Process", "TRL 1-4"):     N(fte=8.5,  cost_myr=6.0,   duration_months=60),
        ("Hardware: Process", "TRL 5-7"):     N(fte=25.0, cost_myr=300.0, duration_months=90),
        ("Algorithm", "TRL 1-4"):            N(fte=3.5,  cost_myr=6.0,   duration_months=18),
        ("Algorithm", "TRL 5-7"):            N(fte=7.5,  cost_myr=30.0,  duration_months=18),
    })
    basf = NormSource(name="BASF", data={
        ("Chemistry", "TRL 1-4"):            N(fte=5.0,  cost_myr=6.0,   duration_months=48),
        ("Chemistry", "TRL 5-7"):            N(fte=27.5, cost_myr=100.0, duration_months=36, technicians=20.0),
        ("Hardware: Mechanical", "TRL 1-4"):  N(fte=5.0,  cost_myr=6.0,   duration_months=48),
        ("Hardware: Mechanical", "TRL 5-7"):  N(fte=27.5, cost_myr=40.0,  duration_months=36, technicians=20.0),
        ("Hardware: Process", "TRL 1-4"):     N(fte=10.0, cost_myr=30.0,  duration_months=48),
        ("Hardware: Process", "TRL 5-7"):     N(fte=60.0, cost_myr=300.0, duration_months=54, technicians=50.0),
        ("Algorithm", "TRL 1-4"):            N(fte=4.5,  cost_myr=8.0,   duration_months=18),
        ("Algorithm", "TRL 5-7"):            N(fte=10.0, cost_myr=24.0,  duration_months=18),
    })
    return [shell, chevron, basf]


def default_norms_config() -> NormsConfig:
    return NormsConfig(
        sources=default_norm_sources(),
        selected_sources=["Shell", "Chevron"],
    )
