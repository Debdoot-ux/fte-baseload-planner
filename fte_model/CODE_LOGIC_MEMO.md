# How This Codebase Works — A Complete Guide

## What This System Does

This system answers one question: **"Given our R&D budget and project portfolio, how many people do we need?"**

You tell it how much money you have, what types of projects you run, how long they take, how much they cost, and how many people work on each one. It tells you how many researchers and developers you need to hire — month by month, year by year.

It does this in three ways:

1. **A Python model** that runs the calculation and powers an interactive dashboard
2. **A standalone Excel workbook** that replicates the same logic in spreadsheet formulas
3. **A normalization workbook** that compares PETRONAS (Malaysia) R&D effort against North American benchmarks

---

## Part 1: The Core Model

### The files

| File | Role |
|---|---|
| `config.py` | Defines all the data structures (what inputs look like, what outputs look like) |
| `defaults.py` | Provides a set of default assumptions (budget, project types, durations, costs) |
| `model.py` | The calculation engine — turns inputs into headcount numbers |

### The inputs

Everything starts with a `ModelConfig` — a bundle of assumptions:

- **Budget**: Total annual R&D spend (default: 400M MYR)
- **Overhead**: What fraction is eaten by admin/facilities before any project gets funded (default: 30%)
- **Archetypes**: The types of projects in your portfolio. There are four:
  - Chemistry (20% of projects)
  - Hardware: Mechanical (30%)
  - Hardware: Process (30%)
  - Algorithm (20%)
- **Pipeline stages**: Projects flow through two stages:
  - TRL 1–4 (early research)
  - TRL 5–7 (late development)
- **Stage mix**: What share of new projects enters at each stage (default: 20% start at TRL 1–4, 80% start directly at TRL 5–7)
- **Conversion rate**: What share of TRL 1–4 completers advance to TRL 5–7 (default: 40%)
- **Intake window**: New projects are spread across the first N months of each year (default: 6 months, meaning Jan–Jun)
- **Utilization**: What fraction of an FTE's time is spent on project work (default: 70% in the baseline)

For each archetype and stage, you also specify:
- **Duration** — how many months a project takes
- **Cost** — total project cost in MYR millions
- **FTE per role** — how many researchers and developers work on one project at any time

### The calculation — five steps

**Step 1: Figure out how much money is available for projects.**

```
Net budget = Total budget × (1 − Overhead %)
```

With defaults: 400M × 0.70 = 280M MYR per year.

**Step 2: Figure out how many new projects you can start each year.**

This is the heart of the model. It uses **cash-flow budgeting**: each year, ongoing projects from prior years consume part of the budget first. Only the leftover funds new projects.

The model tracks every batch of projects (called a "cohort") that started in a given month. Each cohort burns money at a constant monthly rate (= total cost ÷ duration in months) for as long as it is active.

For each year:
1. Add up the cost of all ongoing cohorts from prior years
2. Subtract that from the annual budget
3. Divide the remainder by the cost of starting one new weighted-average project
4. That quotient is the number of new projects for this year

The "cost of one new project" accounts for:
- The portfolio mix (some archetypes are expensive, some are cheap)
- The stage mix (TRL 5–7 projects cost more than TRL 1–4)
- Within-year conversions (a TRL 1–4 project that finishes mid-year and converts to TRL 5–7 starts burning TRL 5–7 budget immediately)
- Intake spreading (projects starting in January burn 12 months of budget that year; projects starting in June burn only 7 months)

In the first year there are no ongoing projects, so the full net budget goes to new starts. In later years, ongoing projects take a bigger bite, leaving less for new starts. This is why the project count drops over time.

**Step 3: Distribute projects across archetypes and stages.**

Each year's project count is split:
- By archetype share (e.g., 20% are Chemistry)
- By stage mix (e.g., 80% enter at TRL 5–7)
- Divided evenly across the intake months

When a TRL 1–4 project finishes, 40% of them convert into new TRL 5–7 projects. These conversion starts are additional — they don't come out of the annual budget allocation (their cost was already accounted for in Step 2).

**Step 4: Simulate the pipeline month by month.**

The model creates a monthly timeline from January of the start year through the end year plus enough tail months for the longest-running projects to finish.

For each month, it counts "active projects" — projects that have started but haven't finished yet. A project that takes 7 months and starts in January is active Jan through Jul (7 months), then gone in August.

This is done with a sliding window: for any given month, active stock = sum of all starts in the window [this month − duration + 1, this month].

**Step 5: Convert active projects into headcount.**

For each month and each archetype/stage:

```
FTE = Active projects × Staff per project ÷ Utilization rate
```

If 10 Chemistry TRL 1–4 projects are active and each needs 3.5 researchers, that's 35 research-FTE before utilization adjustment. At 70% utilization, that becomes 35 ÷ 0.70 = 50 research-FTE (because each person only spends 70% of their time on project work, you need more people).

### The outputs

The model produces:
- **Monthly data**: FTE by archetype, stage, and role for every month in the simulation
- **Annual summary**: For each year — average, minimum, and maximum monthly FTE
- **"Steady state"**: The average FTE in the last intake year (note: this is the planning-horizon peak, not the true mathematical steady state)
- **Cost sensitivity band**: If any archetype has a cost range (min ≠ max), the model runs three times — at midpoint, low cost, and high cost — to show how FTE varies with cost uncertainty

### Validation and guardrails

`config.py` validates inputs on creation:
- Budget cannot be negative (clamped to 0)
- Overhead is clamped to [0%, 100%]
- Intake spread is clamped to [1, 12] months (values above 12 would crash the date arithmetic)
- Archetype shares must sum to 1.0 (±1% tolerance, otherwise a warning is raised)
- Stage mix must sum to 1.0 (±1% tolerance)

### Phase 2 allocation shift

The model supports changing the stage mix partway through the planning horizon. For example:
- 2026–2027: 20% TRL 1–4 / 80% TRL 5–7 (Phase 1)
- 2028–2030: 40% TRL 1–4 / 60% TRL 5–7 (Phase 2)

Set `phase2_start_year` to the year the shift takes effect. Conversion rates stay the same — only the allocation of brand-new projects changes.

---

## Part 2: The Standalone Excel Model

### The file

`build_excel_model.py` generates `FTE_Baseload_Model_Live.xlsx` — a formula-driven Excel workbook that replicates the Python model's logic in spreadsheet formulas.

### Why it exists

The Python model runs in a dashboard. The Excel model exists so that stakeholders can explore the numbers in a familiar tool without needing Python or a running server. You change a yellow cell on the Inputs sheet, and every other sheet updates instantly.

### How it's built

The script runs the Python model once to get project counts (because the cash-flow budget calculation is too granular for Excel formulas — it tracks per-month cohort burn rates). It then writes those counts as constants onto the Budget sheet. Everything else is pure Excel formulas.

### The sheets

**Inputs** — the only sheet you edit. Yellow cells are your assumptions:
- Budget, overhead, start/end year, intake window, utilization
- Stage allocation (% entering at TRL 1–4 vs TRL 5–7) and conversion rate
- Phase 2 allocation shift (optional)
- Portfolio mix (% Chemistry, % HW Mechanical, etc.)
- Per-archetype parameters: duration, cost, Research FTE, Developer FTE for each stage
- Contingency % (optional buffer on top of calculated FTE, set per role)

All inputs are named cells (e.g., `Budget`, `Overhead`, `Chem_E_Dur`) so formulas read like English.

**Budget** — one row per intake year (2026–2030). Shows the number of new projects per year. These are constants pre-computed by the Python cash-flow model. Changing budget or overhead on the Inputs sheet does NOT update these counts (you'd need to re-run the Python script). This is a known limitation.

**Engine** — 156 rows (13 years × 12 months). Each archetype gets a block of 9 columns:

| Column | What it calculates |
|---|---|
| Early Starts/mo | New TRL 1–4 projects starting this month = Projects × Share × Allocation ÷ Intake months |
| Early Active | Sliding window sum of starts over the past `duration` months |
| Late Conv Starts | TRL 1–4 completers converting to TRL 5–7 = starts from `duration` months ago × conversion rate |
| Late Direct Starts | New TRL 5–7 projects starting directly (same formula pattern as early starts) |
| Late Total Starts | Conversion + Direct |
| Late Active | Sliding window sum of late total starts |
| Research FTE | (Early Active × Research per early + Late Active × Research per late) ÷ Utilization |
| Developer FTE | Same pattern with Developer parameters |
| Total FTE | Research + Developer |

After the four archetype blocks come:
- **Totals**: Sum of Research/Developer/Total across all archetypes
- **Adjusted Totals**: Totals × (1 + Contingency %) — the final planning numbers

The Engine formulas are phase-aware: they check whether the current year is before or after the Phase 2 shift year and use the corresponding allocation.

**Output** — annual summary. For each year (2026–2030), shows min/max/avg monthly FTE from the Engine, both base and adjusted. Uses MINIFS/MAXIFS/AVERAGEIFS formulas filtering by year.

**Glossary** — plain-English explanation of how the model works, what the terms mean, and what assumptions are baked in.

### What updates automatically vs. what doesn't

Changes on Inputs that **do** propagate instantly:
- Portfolio shares, FTE per project, duration, utilization, contingency, stage allocation, conversion rate, Phase 2 settings

Changes that **don't** propagate (require re-running the Python script):
- Budget, overhead, archetype costs (these affect project counts, which are frozen constants on the Budget sheet)

This means the Excel model answers: *"Given these project counts, how many people?"* — not *"Given this budget, how many people?"* The Streamlit dashboard answers the full question because it re-runs the Python model on every change.

---

## Part 3: The Normalization Analysis

### The file

`build_normalization_excel.py` generates `MY_vs_NA_Normalization.xlsx` — a workbook that compares PETRONAS R&D effort against North American benchmarks.

### The problem it solves

Comparing raw budgets across regions is misleading:
- MY Chemistry budget: 78M MYR
- NA Chemistry budget: 78M USD

These look "equal" but 78M MYR is only ~19.5M USD. And even after currency conversion, a researcher in Malaysia costs 80K USD/year vs 200K USD/year in North America. Cheaper people ≠ less effort.

### The metric: person-years

**Person-years = Budget ÷ Local fully loaded FTE cost**

This converts money into a universal unit: how many years of full-time human work does this budget represent? It strips out both currency and wage differences.

### The inputs

| Parameter | Value | Source |
|---|---|---|
| FX rate | 4.0 MYR per USD | Market rate |
| NA FTE cost | 200,000 USD/year | Industry benchmark |
| MY FTE cost | 80,000 USD/year | PETRONAS data |
| MY total budget | 400M MYR | Constrained budget scenario |
| MY overhead | 52% | PETRONAS data |
| Archetype shares | Chemistry 40%, HW Mech 18%, HW Process 25%, Algorithm 17% | Portfolio analysis |
| NA budgets | Chemistry 78M, HW Mech 48M, HW Process 50M, Algorithm 48M USD | Outside-in estimates |

### The calculation (portfolio level)

**Step 1**: Split MY budget by archetype, convert to USD.
```
MY net budget = 400M × (1 − 52%) = 192M MYR
MY Chemistry = 192M × 40% = 76.8M MYR = 19.2M USD
```

**Step 2**: Divide each budget by local FTE cost to get raw person-years.
```
MY Chemistry PY = 19,200,000 ÷ 80,000 = 240 person-years
NA Chemistry PY = 78,000,000 ÷ 200,000 = 390 person-years
```

**Step 3**: Compare. Optional productivity factors can adjust for differences in lab quality/support.

Result: MY deploys ~30–40% fewer person-years than NA across most archetypes.

### The calculation (project level)

The workbook also compares effort at the individual project level, using data from actual PETRONAS reference projects (Fgo, Garcinia, Elektra, ASAT, etc.) and outside-in estimates for Shell and BASF.

For each archetype and TRL stage, it shows:
- Duration (months), team size (FTE), and person-years (= FTE × duration ÷ 12)
- Cost per person-year (how capital-intensive the work is)
- Shell/MY and BASF/MY multipliers (how many times more effort the benchmarks use)

Shell and BASF cost-per-PY values are calculated as Excel formulas referencing the FX rate on the Inputs sheet, so changing FX automatically updates these comparisons.

### The sheets

| Sheet | What it shows |
|---|---|
| Inputs | All editable assumptions (yellow cells). Change FX, salaries, overhead, shares. |
| Calculation | Portfolio-level normalization in three steps. All formulas referencing Inputs. |
| Project Comparison | Per-project person-years: PETRONAS vs Shell vs BASF. Static benchmark data + FX-linked cost/PY. |
| Reference Data | Raw project-by-project PETRONAS data behind the MY averages. |
| Methodology | Plain-English explanation of what the workbook does, how to read it, key caveats. |
| Sensitivity | Tests what happens to the portfolio ratio if NA budgets are 3×, 4×, or 5× MY budget. |

### What this means

After removing currency and wage effects, MY puts in fewer person-years than NA across all archetypes. There are only two explanations:

1. **MY is more productive** — achieving comparable R&D output with fewer people
2. **MY is doing less work** — smaller scope, fewer projects, or less complex research

Budget data alone cannot tell us which. The defensible statement is: *"Normalized for wages, MY R&D programs deploy 30–40% fewer person-years than NA benchmarks across most archetypes."*

### Key caveat: circularity risk

Three of four NA budgets are numerically identical to MY budgets but in USD vs MYR (e.g., Chemistry: 78M USD vs 78M MYR). If the outside-in estimates were derived from MY numbers by applying the FX rate, the comparison is circular. HW Process (50M vs 72M) is the only archetype where budgets clearly differ independently. The NA estimates should be confirmed as independently sourced.

---

## Part 4: The Scenario System

### The files

| File | Role |
|---|---|
| `scenario_parser.py` | Reads scenario definitions from uploaded Excel files |
| `scenario_engine.py` | Runs the model for each scenario and produces comparison outputs |

### How scenario parsing works

Users can upload an Excel file with a "scenarios" sheet and an "assumptions" sheet. The parser:

1. **Identifies sheets** by scanning for keyword density (e.g., a row containing "scenario", "budget", "overhead", "split" is likely the scenario header)
2. **Maps columns** by matching header text to known patterns (regex-based)
3. **Reads rows** — each row is a scenario with budget, overhead, stage mix, conversion rate, and archetype shares
4. **Inherits missing values** — if a scenario row is blank for some field, it inherits from the first (baseline) scenario
5. **Parses assumptions** — archetype definitions with duration, FTE, and cost per stage
6. **Matches archetype names** between sheets using a 4-pass strategy:
   - Exact match (case-insensitive)
   - Known abbreviations (HP → Hardware Process, AI → Algorithm)
   - Substring match
   - Positional fallback (with a warning)

The parser also normalizes data: values > 1.0 for overhead, conversion rates, and shares are treated as percentages and divided by 100. Phase labels like "> TRL 4" or "TRL 5 - 7" are normalized to "TRL 5-7".

### How scenario comparison works

`scenario_engine.py` takes a list of (name, config) pairs, runs `run_model()` on each, and produces:
- A summary table (one row per scenario: FTE, budget, peak headcount, etc.)
- Per-scenario Excel sheets with annual summaries and archetype parameters
- A downloadable comparison Excel workbook

---

## Part 5: The Dashboard

### The file

`app.py` — a Streamlit web application.

### What it does

The dashboard is the interactive front-end. It lets users:

1. **Configure scenarios** — adjust budget, overhead, archetypes, stage mix, conversion rates, cost ranges, utilization, and contingency through sliders and input fields
2. **View results** — monthly FTE charts, annual summary tables, archetype breakdowns, role splits
3. **Compare scenarios** — side-by-side comparison across multiple configurations
4. **Upload Excel inputs** — parse scenario files using `scenario_parser.py`
5. **Download results** — export to Excel using `scenario_engine.py`

Every time you change a slider, the Python model re-runs from scratch. This means the dashboard always shows the full budget→projects→FTE chain, unlike the Excel model where project counts are frozen.

---

## Part 6: How Everything Connects

```
User inputs (dashboard sliders OR Excel upload)
    │
    ▼
ModelConfig (config.py)
    │
    ▼
model.py: _compute_yearly_projects()     ← Cash-flow budget → project counts
    │
    ▼
model.py: _run_archetype()               ← Project counts → monthly active stock → FTE
    │
    ▼
model.py: run_model()                    ← Aggregates into ModelResult
    │
    ├──▶ app.py (Streamlit dashboard)    ← Interactive charts and tables
    ├──▶ scenario_engine.py              ← Multi-scenario comparison
    └──▶ build_excel_model.py            ← Standalone Excel workbook
```

The normalization workbook (`build_normalization_excel.py`) is independent — it doesn't import the model. It's a standalone analysis with its own hardcoded data.

---

## Part 7: Key Assumptions and Limitations

### Assumptions baked into the model

1. **Constant annual budget** — the same amount every year across the planning horizon. No year-over-year growth or decline.
2. **Linear pipeline** — projects flow TRL 1–4 → TRL 5–7. No branching, looping, or skipping back.
3. **No mid-stage cancellation** — once a project starts, it runs to completion.
4. **Instant staffing** — projects start at full team size immediately. No ramp-up period (though the code supports `ramp_months`, it defaults to 0).
5. **Even intake spreading** — new projects are distributed uniformly across the intake window. No seasonality.
6. **No economies of scale** — cost per project is constant regardless of how many you run.
7. **Two roles only** — Researcher and Developer. The model could support more, but the Excel workbook is hardcoded to two.
8. **Four archetypes** — Chemistry, HW Mechanical, HW Process, Algorithm. Adding more requires rebuilding the Excel workbook.

### Known limitations

1. **Excel Budget sheet is a snapshot** — project counts come from the Python model and are written as constants. Changing budget/overhead/costs on the Inputs sheet does NOT recalculate project counts. Only the FTE-per-project calculations update in real time.
2. **"Steady state" is approximate** — the model reports FTE for the last intake year as "steady state." For long-duration archetypes like Hardware: Process (24 + 75 = 99 months), the pipeline doesn't fully stabilize until several years after the last intake. The reported figure is a planning-horizon metric, not the true mathematical steady state.
3. **Cascading within-year conversions are incomplete** — if a TRL 1–4 project finishes mid-year, converts to TRL 5–7, and that TRL 5–7 project also finishes within the same year, the second conversion is not tracked. In practice this only matters for Algorithm (TRL 5–7 = 9 months) and the budget impact is small.
4. **Cost sensitivity triples computation** — when any archetype has a cost range (min ≠ max), the model runs three times (low, mid, high). With 5+ scenarios, this means 15+ model runs per interaction.

---

## Glossary

| Term | Meaning |
|---|---|
| **FTE** | Full-Time Equivalent. 1 FTE = one person working full time. |
| **TRL** | Technology Readiness Level. A 1-to-9 scale of technology maturity. |
| **Archetype** | A type of R&D project (e.g., Chemistry, Hardware: Mechanical). |
| **Stage mix** | The allocation of new projects across pipeline stages (e.g., 20% TRL 1–4, 80% TRL 5–7). |
| **Conversion rate** | The percentage of projects finishing one stage that advance to the next. |
| **Intake window** | The number of months per year during which new projects start (e.g., 6 = Jan–Jun). |
| **Utilization** | The fraction of an FTE's time spent on project work. At 70%, you need 1/0.70 = 1.43 people per "project FTE." |
| **Contingency** | A percentage buffer added on top of calculated FTE to cover uncertainty, attrition, or estimation error. |
| **Cohort** | A batch of projects that all started in the same month. The model tracks cohorts to calculate ongoing costs. |
| **Active stock** | The number of projects running in a given month. Computed as a sliding window over starts. |
| **Person-year** | One person working full time for one year. Used in the normalization analysis to compare effort across regions. |
| **Cash-flow budgeting** | The approach where each year's budget must cover ongoing project costs before funding new starts. |
| **Commitment budgeting** | An alternative approach (selectable via `budget_mode`) where the full lifecycle cost is committed upfront and the same number of projects starts every year. |
| **Phase 2** | An optional shift in stage allocation partway through the planning horizon. |
| **Net budget** | Total budget minus overhead. The money that actually funds projects. |
