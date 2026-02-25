"""
FTE Baseload Calculation Engine
Single-scenario model: user inputs drive one calculation.
Annual summary shows within-year range (min/max monthly FTE).
"""

from __future__ import annotations

from typing import Dict, List

import pandas as pd

from config import Archetype, ModelConfig, ModelResult, StageParams


def _available_budget(cfg: ModelConfig) -> float:
    return cfg.total_budget_m * (1 - cfg.overhead_pct)


# ---------------------------------------------------------------------------
# Cost calculation (generalized for N stages)
# ---------------------------------------------------------------------------

def _expected_cost_from_stage(
    arch: Archetype,
    stages: List[str],
    conv_rates: Dict[str, float],
    start_idx: int,
) -> float:
    """Expected cost for a project entering at stages[start_idx], including
    probabilistic conversion to later stages."""
    if start_idx >= len(stages):
        return 0.0
    sname = stages[start_idx]
    if sname not in arch.stages:
        return 0.0
    cost = arch.stages[sname].cost_millions
    if start_idx < len(stages) - 1:
        conv = conv_rates.get(sname, 0.0)
        if conv > 0:
            cost += conv * _expected_cost_from_stage(
                arch, stages, conv_rates, start_idx + 1
            )
    return cost


def _weighted_cost_per_project(cfg: ModelConfig) -> float:
    """Portfolio-weighted average cost per project across all entry points."""
    stages = cfg.pipeline_stages
    total = 0.0
    for arch in cfg.archetypes:
        arch_cost = 0.0
        for i, sname in enumerate(stages):
            mix = cfg.stage_mix.get(sname, 0.0)
            if mix > 0 and sname in arch.stages:
                arch_cost += mix * _expected_cost_from_stage(
                    arch, stages, cfg.stage_conversion_rates, i
                )
        total += arch.portfolio_share * arch_cost
    return total


def _projects_per_year(cfg: ModelConfig) -> float:
    budget = _available_budget(cfg)
    wc = _weighted_cost_per_project(cfg)
    if wc <= 0:
        return 0.0
    return budget / wc


# ---------------------------------------------------------------------------
# Active stock (with optional ramp)
# ---------------------------------------------------------------------------

def _active_stock(
    starts: pd.Series,
    duration: int,
    idx: pd.DatetimeIndex,
    ramp_months: int = 0,
) -> pd.Series:
    active = pd.Series(0.0, index=idx)
    if duration <= 0:
        return active
    if ramp_months <= 0:
        for dt, n in starts.items():
            if n < 1e-9:
                continue
            end_dt = dt + pd.DateOffset(months=duration)
            mask = (idx >= dt) & (idx < end_dt)
            active.loc[mask] += n
    else:
        for dt, n in starts.items():
            if n < 1e-9:
                continue
            for m in range(duration):
                m_dt = dt + pd.DateOffset(months=m)
                if m_dt in active.index:
                    factor = min(1.0, (m + 1) / ramp_months)
                    active.loc[m_dt] += n * factor
    return active


# ---------------------------------------------------------------------------
# Pipeline runner (N-stage, per-archetype)
# ---------------------------------------------------------------------------

def _run_archetype(
    cfg: ModelConfig,
    arch: Archetype,
    arch_projects: float,
    idx: pd.DatetimeIndex,
    utilization: float,
    records: List[dict],
) -> None:
    stages = cfg.pipeline_stages
    prev_completions: pd.Series | None = None

    for i, sname in enumerate(stages):
        if sname not in arch.stages:
            prev_completions = None
            continue

        sp = arch.stages[sname]
        dur = sp.duration_months
        fte_r = sp.fte_research
        fte_d = sp.fte_developer

        starts = pd.Series(0.0, index=idx)

        direct_n = arch_projects * cfg.stage_mix.get(sname, 0.0)
        if direct_n > 0:
            monthly_n = direct_n / max(cfg.intake_spread_months, 1)
            for y in range(cfg.start_year, cfg.end_year + 1):
                for m in range(1, cfg.intake_spread_months + 1):
                    ts = pd.Timestamp(f"{y}-{m:02d}-01")
                    if ts in starts.index:
                        starts.loc[ts] += monthly_n

        if prev_completions is not None and i > 0:
            prev_sname = stages[i - 1]
            conv = cfg.stage_conversion_rates.get(prev_sname, 0.0)
            if conv > 0:
                starts = starts.add(prev_completions * conv, fill_value=0.0)

        active = _active_stock(starts, dur, idx, cfg.ramp_months)

        this_completions = pd.Series(0.0, index=idx)
        for dt, n in starts.items():
            if n < 1e-9:
                continue
            comp_dt = dt + pd.DateOffset(months=dur)
            if comp_dt in this_completions.index:
                this_completions.loc[comp_dt] += n
        prev_completions = this_completions

        for dt in idx:
            n = active.loc[dt]
            if n < 1e-9:
                continue
            records.append({
                "month": dt,
                "year": dt.year,
                "archetype": arch.name,
                "stage": sname,
                "effective_projects": n,
                "fte_research": n * fte_r / utilization,
                "fte_developer": n * fte_d / utilization,
                "fte_total": n * (fte_r + fte_d) / utilization,
            })


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def run(cfg: ModelConfig) -> Dict:
    total_proj = _projects_per_year(cfg)

    tail_months = 0
    for a in cfg.archetypes:
        arch_dur = sum(
            sp.duration_months
            for sn, sp in a.stages.items()
            if sn in cfg.pipeline_stages
        )
        tail_months = max(tail_months, arch_dur)

    start_date = pd.Timestamp(f"{cfg.start_year}-01-01")
    end_date = pd.Timestamp(f"{cfg.end_year}-12-01") + pd.DateOffset(months=tail_months)
    idx = pd.date_range(start_date, end_date, freq="MS")

    utilization = max(cfg.utilization_rate, 0.01)
    records: List[dict] = []

    for arch in cfg.archetypes:
        arch_projects = total_proj * arch.portfolio_share
        _run_archetype(cfg, arch, arch_projects, idx, utilization, records)

    monthly = pd.DataFrame(records)
    if monthly.empty:
        monthly = pd.DataFrame(
            columns=["month", "year", "archetype", "stage",
                      "effective_projects", "fte_research", "fte_developer", "fte_total"]
        )

    return {"monthly": monthly, "projects_per_year": total_proj}


def run_model(cfg: ModelConfig) -> ModelResult:
    res = run(cfg)
    monthly = res["monthly"]

    annual = _build_annual_summary(monthly, cfg)

    ss_avg, ss_min, ss_max = _steady_state(monthly, cfg)

    return ModelResult(
        monthly=monthly,
        annual_summary=annual,
        steady_state_avg=ss_avg,
        steady_state_min_month=ss_min,
        steady_state_max_month=ss_max,
        projects_per_year=res["projects_per_year"],
    )


def _build_annual_summary(monthly, cfg):
    rows = []
    if monthly.empty:
        return pd.DataFrame(rows)

    for yr in range(cfg.start_year, cfg.end_year + 1):
        yr_data = monthly[monthly["year"] == yr]
        if yr_data.empty:
            rows.append({
                "Year": yr,
                "Avg monthly FTE": 0,
                "Min monthly FTE": 0,
                "Max monthly FTE": 0,
                "Avg Research FTE": 0,
                "Avg Developer FTE": 0,
            })
            continue

        monthly_totals = yr_data.groupby("month").agg(
            total=("fte_total", "sum"),
            research=("fte_research", "sum"),
            developer=("fte_developer", "sum"),
        )

        rows.append({
            "Year": yr,
            "Avg monthly FTE": round(monthly_totals["total"].mean(), 1),
            "Min monthly FTE": round(monthly_totals["total"].min(), 1),
            "Max monthly FTE": round(monthly_totals["total"].max(), 1),
            "Avg Research FTE": round(monthly_totals["research"].mean(), 1),
            "Avg Developer FTE": round(monthly_totals["developer"].mean(), 1),
        })

    return pd.DataFrame(rows)


def _steady_state(monthly, cfg):
    """Return (avg, min_month, max_month) FTE for the last intake year."""
    if monthly.empty:
        return 0.0, 0.0, 0.0
    last_year = monthly[monthly["year"] == cfg.end_year]
    if last_year.empty:
        last_year = monthly[monthly["year"] == cfg.end_year - 1]
    if last_year.empty:
        return 0.0, 0.0, 0.0

    monthly_totals = last_year.groupby("month")["fte_total"].sum()
    return (
        monthly_totals.mean(),
        monthly_totals.min(),
        monthly_totals.max(),
    )
