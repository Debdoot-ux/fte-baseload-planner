# FTE Baseload Planning Tool

An R&D workforce planning model that answers: **"Given our R&D budget and project portfolio, how many researchers and developers do we need?"**

## Setup

```bash
pip install -r requirements.txt
```

## How to Run

### Streamlit App (primary UI)

```bash
cd fte_model
streamlit run app.py
```

Configure scenarios manually — budget, overhead, project types, stage durations, team sizes — then calculate headcount.

### Excel Model Builder

```bash
cd fte_model
python build_excel_model.py
```

Generates `FTE_Baseload_Model_Live.xlsx` — a formula-driven Excel workbook with editable inputs (yellow cells) and auto-calculated outputs. The year range is fixed to 2026–2030.

### Tests

```bash
cd fte_model
python e2e_test.py
```

## Project Structure

| File | Purpose |
|------|---------|
| `config.py` | Data classes: `StageParams`, `Archetype`, `ModelConfig`, `ModelResult` |
| `defaults.py` | Baseline `ModelConfig` with 4 archetypes and default parameters |
| `model.py` | Core calculation engine: cash-flow budgeting, cohort tracking, monthly FTE |
| `scenario_engine.py` | Multi-scenario runner, comparison summary, comparison Excel export |
| `app.py` | Streamlit UI: multi-scenario config, results, comparison, Excel export |
| `build_excel_model.py` | Generates a standalone formula-driven Excel model |
| `e2e_test.py` | End-to-end tests covering all code paths |

## How the Model Works

1. **Budget** — Total R&D budget minus overhead = net project budget
2. **Projects** — Net budget divided by weighted cost per project = number of projects (cash-flow mode deducts ongoing costs first)
3. **Pipeline** — Projects are distributed across archetypes and stages, tracked month by month
4. **FTE** — Active projects × team size per role = monthly headcount
5. **Steady state** — Once new starts roughly equal completions, headcount stabilizes
