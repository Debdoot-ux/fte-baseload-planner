"""
FTE Baseload Calculation Engine
Cash-flow budget model: annual budget covers ongoing + new project costs.
Annual summary shows within-year range (min/max monthly FTE).
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import Dict, List, Tuple

import pandas as pd

from config import Archetype, ModelConfig, ModelResult, StageParams


def _available_budget(cfg: ModelConfig) -> float:
    return cfg.total_budget_m * (1 - cfg.overhead_pct)


def _get_stage_mix(cfg: ModelConfig, year: int) -> Dict[str, float]:
    """Return the stage mix applicable for *year*, respecting Phase 2 override."""
    if cfg.phase2_start_year > 0 and cfg.stage_mix_phase2 and year >= cfg.phase2_start_year:
        return cfg.stage_mix_phase2
    return cfg.stage_mix


# ---------------------------------------------------------------------------
# Cost helpers (kept for backward-compat; used by app.py assumption register)
# ---------------------------------------------------------------------------

def _expected_cost_from_stage(
    arch: Archetype,
    stages: List[str],
    conv_rates: Dict[str, float],
    start_idx: int,
) -> float:
    """Expected lifecycle cost for a project entering at stages[start_idx],
    including probabilistic conversion to later stages."""
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


def weighted_cost_per_project(cfg: ModelConfig, mix: Dict[str, float] | None = None) -> float:
    """Portfolio-weighted average lifecycle cost per project."""
    if mix is None:
        mix = cfg.stage_mix
    stages = cfg.pipeline_stages
    total = 0.0
    for arch in cfg.archetypes:
        arch_cost = 0.0
        for i, sname in enumerate(stages):
            mix_val = mix.get(sname, 0.0)
            if mix_val > 0 and sname in arch.stages:
                arch_cost += mix_val * _expected_cost_from_stage(
                    arch, stages, cfg.stage_conversion_rates, i
                )
        total += arch.portfolio_share * arch_cost
    return total


def projects_per_year(cfg: ModelConfig) -> float:
    """Commitment-model project count (for display only)."""
    budget = _available_budget(cfg)
    wc = weighted_cost_per_project(cfg)
    if wc <= 0:
        return 0.0
    return budget / wc


def _projects_for_year(cfg: ModelConfig, year: int) -> float:
    """Commitment-model project count for a specific year (for display only)."""
    mix = _get_stage_mix(cfg, year)
    budget = _available_budget(cfg)
    wc = weighted_cost_per_project(cfg, mix)
    if wc <= 0:
        return 0.0
    return budget / wc


# ---------------------------------------------------------------------------
# Cash-flow budget model
# ---------------------------------------------------------------------------

@dataclass
class _Cohort:
    """A batch of projects that started in the same month."""
    start_month_offset: int   # months from model start (0 = Jan of start_year)
    count: float
    monthly_burn: float       # cost_millions / duration_months
    duration: int             # months
    stage_idx: int            # index into pipeline_stages
    arch_idx: int             # index into archetypes


def _months_active_in_year(cohort: _Cohort, year_offset: int) -> int:
    """How many months of `year_offset` (0-indexed from start_year) this cohort
    is active during. year_offset 0 = start_year."""
    year_first = year_offset * 12
    year_last = year_first + 11
    cohort_first = cohort.start_month_offset
    cohort_last = cohort_first + cohort.duration - 1
    overlap_first = max(cohort_first, year_first)
    overlap_last = min(cohort_last, year_last)
    return max(0, overlap_last - overlap_first + 1)


def _partial_year_cost_per_new_project(
    cfg: ModelConfig,
    year: int,
    year_offset: int,
) -> float:
    """Cost consumed in `year` by starting 1 weighted-average new project,
    including within-year conversion costs."""
    mix = _get_stage_mix(cfg, year)
    stages = cfg.pipeline_stages
    intake = max(cfg.intake_spread_months, 1)
    year_first_abs = year_offset * 12

    total_cost = 0.0
    for arch in cfg.archetypes:
        arch_cost = 0.0
        for m in range(intake):
            abs_start = year_first_abs + m
            for si, sname in enumerate(stages):
                if sname not in arch.stages:
                    continue
                sp = arch.stages[sname]
                mix_val = mix.get(sname, 0.0)
                if mix_val <= 0:
                    continue

                burn = sp.cost_millions / max(sp.duration_months, 1)
                months_in_year = min(12 - m, sp.duration_months)
                direct_cost = mix_val * burn * months_in_year
                arch_cost += direct_cost / intake

                if si < len(stages) - 1:
                    conv = cfg.stage_conversion_rates.get(sname, 0.0)
                    comp_month_in_year = m + sp.duration_months
                    if conv > 0 and comp_month_in_year < 12:
                        next_sname = stages[si + 1]
                        if next_sname in arch.stages:
                            next_sp = arch.stages[next_sname]
                            next_burn = next_sp.cost_millions / max(next_sp.duration_months, 1)
                            next_months = min(12 - comp_month_in_year, next_sp.duration_months)
                            conv_cost = mix_val * conv * next_burn * next_months
                            arch_cost += conv_cost / intake

        total_cost += arch.portfolio_share * arch_cost
    return total_cost


def _compute_yearly_projects(cfg: ModelConfig) -> Dict[int, float]:
    """Compute new project counts year-by-year under cash-flow budgeting."""
    budget = _available_budget(cfg)
    stages = cfg.pipeline_stages
    intake = max(cfg.intake_spread_months, 1)

    cohorts: List[_Cohort] = []
    yearly: Dict[int, float] = {}

    for y_off, year in enumerate(range(cfg.start_year, cfg.end_year + 1)):
        mix = _get_stage_mix(cfg, year)

        new_conv_cohorts: List[_Cohort] = []
        for coh in cohorts:
            if coh.stage_idx >= len(stages) - 1:
                continue
            sname = stages[coh.stage_idx]
            conv = cfg.stage_conversion_rates.get(sname, 0.0)
            if conv <= 0:
                continue

            comp_offset = coh.start_month_offset + coh.duration
            year_first = y_off * 12
            year_last = year_first + 11

            if year_first <= comp_offset <= year_last:
                next_si = coh.stage_idx + 1
                arch = cfg.archetypes[coh.arch_idx]
                next_sname = stages[next_si]
                if next_sname in arch.stages:
                    next_sp = arch.stages[next_sname]
                    new_conv_cohorts.append(_Cohort(
                        start_month_offset=comp_offset,
                        count=coh.count * conv,
                        monthly_burn=next_sp.cost_millions / max(next_sp.duration_months, 1),
                        duration=next_sp.duration_months,
                        stage_idx=next_si,
                        arch_idx=coh.arch_idx,
                    ))
        cohorts.extend(new_conv_cohorts)

        ongoing_cost = 0.0
        for coh in cohorts:
            m_active = _months_active_in_year(coh, y_off)
            ongoing_cost += coh.count * coh.monthly_burn * m_active

        available = max(0.0, budget - ongoing_cost)
        partial_cost = _partial_year_cost_per_new_project(cfg, year, y_off)
        n_new = available / partial_cost if partial_cost > 0 else 0.0
        yearly[year] = n_new

        for ai, arch in enumerate(cfg.archetypes):
            for si, sname in enumerate(stages):
                if sname not in arch.stages:
                    continue
                sp = arch.stages[sname]
                mix_val = mix.get(sname, 0.0)
                if mix_val <= 0:
                    continue

                count_per_month = n_new * arch.portfolio_share * mix_val / intake
                if count_per_month < 1e-12:
                    continue

                burn = sp.cost_millions / max(sp.duration_months, 1)
                for m in range(intake):
                    abs_month = y_off * 12 + m
                    cohorts.append(_Cohort(
                        start_month_offset=abs_month,
                        count=count_per_month,
                        monthly_burn=burn,
                        duration=sp.duration_months,
                        stage_idx=si,
                        arch_idx=ai,
                    ))

        year_first = y_off * 12
        year_last = year_first + 11
        for ai, arch in enumerate(cfg.archetypes):
            for si, sname in enumerate(stages):
                if si >= len(stages) - 1:
                    continue
                if sname not in arch.stages:
                    continue
                sp = arch.stages[sname]
                conv = cfg.stage_conversion_rates.get(sname, 0.0)
                if conv <= 0:
                    continue
                mix_val = mix.get(sname, 0.0)
                if mix_val <= 0:
                    continue

                next_si = si + 1
                next_sname = stages[next_si]
                if next_sname not in arch.stages:
                    continue
                next_sp = arch.stages[next_sname]

                count_per_month = n_new * arch.portfolio_share * mix_val / intake
                for m in range(intake):
                    comp_offset = y_off * 12 + m + sp.duration_months
                    if year_first <= comp_offset <= year_last:
                        cohorts.append(_Cohort(
                            start_month_offset=comp_offset,
                            count=count_per_month * conv,
                            monthly_burn=next_sp.cost_millions / max(next_sp.duration_months, 1),
                            duration=next_sp.duration_months,
                            stage_idx=next_si,
                            arch_idx=ai,
                        ))

    return yearly


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
    idx: pd.DatetimeIndex,
    utilization: float,
    records: List[dict],
    yearly_projects: Dict[int, float],
) -> None:
    stages = cfg.pipeline_stages
    prev_completions: pd.Series | None = None

    for i, sname in enumerate(stages):
        if sname not in arch.stages:
            prev_completions = None
            continue

        sp = arch.stages[sname]
        dur = sp.duration_months
        fte_roles = sp.fte_per_role
        fte_total_per_project = sum(fte_roles.values())

        starts = pd.Series(0.0, index=idx)

        for y in range(cfg.start_year, cfg.end_year + 1):
            mix_y = _get_stage_mix(cfg, y)
            proj_y = yearly_projects.get(y, 0.0) * arch.portfolio_share
            direct_n = proj_y * mix_y.get(sname, 0.0)
            if direct_n > 0:
                monthly_n = direct_n / max(cfg.intake_spread_months, 1)
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
            record = {
                "month": dt,
                "year": dt.year,
                "archetype": arch.name,
                "stage": sname,
                "effective_projects": n,
                "fte_total": n * fte_total_per_project / utilization,
            }
            for role_name, fte_count in fte_roles.items():
                record[f"fte_{role_name}"] = n * fte_count / utilization
            records.append(record)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def _commitment_yearly_projects(cfg: ModelConfig) -> Dict[int, float]:
    """Compute per-year project counts under the commitment model."""
    yearly: Dict[int, float] = {}
    for year in range(cfg.start_year, cfg.end_year + 1):
        yearly[year] = _projects_for_year(cfg, year)
    return yearly


def run(cfg: ModelConfig) -> Dict:
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

    if cfg.budget_mode == "commitment":
        yearly_projects = _commitment_yearly_projects(cfg)
    else:
        yearly_projects = _compute_yearly_projects(cfg)

    for arch in cfg.archetypes:
        _run_archetype(cfg, arch, idx, utilization, records, yearly_projects)

    monthly = pd.DataFrame(records)
    if monthly.empty:
        base_cols = ["month", "year", "archetype", "stage", "effective_projects", "fte_total"]
        role_cols = [f"fte_{r}" for r in cfg.all_roles]
        monthly = pd.DataFrame(columns=base_cols + role_cols)

    return {
        "monthly": monthly,
        "projects_per_year": yearly_projects.get(cfg.end_year, 0.0),
        "yearly_projects": yearly_projects,
    }


def run_model(cfg: ModelConfig) -> ModelResult:
    res = run(cfg)
    monthly = res["monthly"]

    annual = _build_annual_summary(monthly, cfg)
    ss_avg, ss_min, ss_max = _steady_state(monthly, cfg)

    result = ModelResult(
        monthly=monthly,
        annual_summary=annual,
        steady_state_avg=ss_avg,
        steady_state_min_month=ss_min,
        steady_state_max_month=ss_max,
        projects_per_year=res["projects_per_year"],
        yearly_projects=res["yearly_projects"],
    )

    # Cost sensitivity: run at low/high cost if any stage has a range
    has_range = any(
        abs(sp.cost_max - sp.cost_min) > 0.001
        for arch in cfg.archetypes
        for sp in arch.stages.values()
    )
    if has_range:
        cfg_low = copy.deepcopy(cfg)
        for arch in cfg_low.archetypes:
            for sp in arch.stages.values():
                sp.cost_max = sp.cost_min
        res_low = run(cfg_low)
        low_annual = _build_annual_summary(res_low["monthly"], cfg_low)
        low_ss, _, _ = _steady_state(res_low["monthly"], cfg_low)

        cfg_high = copy.deepcopy(cfg)
        for arch in cfg_high.archetypes:
            for sp in arch.stages.values():
                sp.cost_min = sp.cost_max
        res_high = run(cfg_high)
        high_annual = _build_annual_summary(res_high["monthly"], cfg_high)
        high_ss, _, _ = _steady_state(res_high["monthly"], cfg_high)

        result.cost_low_ss_avg = low_ss
        result.cost_high_ss_avg = high_ss
        result.cost_low_annual = low_annual
        result.cost_high_annual = high_annual

    return result


def _build_annual_summary(monthly, cfg):
    rows = []
    if monthly.empty:
        return pd.DataFrame(rows)

    all_roles = cfg.all_roles
    role_cols = [f"fte_{role}" for role in all_roles
                 if f"fte_{role}" in monthly.columns]

    for yr in range(cfg.start_year, cfg.end_year + 1):
        yr_data = monthly[monthly["year"] == yr]
        if yr_data.empty:
            row = {
                "Year": yr,
                "Avg monthly FTE": 0,
                "Min monthly FTE": 0,
                "Max monthly FTE": 0,
            }
            for role in all_roles:
                row[f"Avg {role} FTE"] = 0
            rows.append(row)
            continue

        agg_dict = {"total": ("fte_total", "sum")}
        for rc in role_cols:
            agg_dict[rc] = (rc, "sum")

        monthly_totals = yr_data.groupby("month").agg(**agg_dict)

        row = {
            "Year": yr,
            "Avg monthly FTE": round(monthly_totals["total"].mean(), 1),
            "Min monthly FTE": round(monthly_totals["total"].min(), 1),
            "Max monthly FTE": round(monthly_totals["total"].max(), 1),
        }
        for role in all_roles:
            col = f"fte_{role}"
            if col in monthly_totals.columns:
                row[f"Avg {role} FTE"] = round(monthly_totals[col].mean(), 1)
            else:
                row[f"Avg {role} FTE"] = 0
        rows.append(row)

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
