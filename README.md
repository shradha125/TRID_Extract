# TRID Urban-Freight Microsimulation Review — Toolkit

## What this is
A reproducible pipeline that turns a TRID-exported `.docx` of freight-simulation papers into a reviewable Excel matrix of research findings, method characteristics, policy focus, geography, and inferred ML/AI gaps.

Run on your two TRID exports (120 publications + 2 projects), the pipeline produced **`trid_review_matrix.xlsx`** — 119 papers coded across 17 columns, with filters and a summary sheet.

## Files

| File | Purpose |
|---|---|
| `parse_trid.py` | Parses TRID `.docx` into structured records (title, authors, year, abstract, venue, keywords). Handles quirks: flush-left author lines, missing fields on old records, pre-2000 years. |
| `records_clean.json` | 119 cleaned records from your two input files. |
| `extractions.py` | Hand-curated research-content extraction (17-column schema) for the current 119 records. |
| `extract_via_api.py` | Reusable Anthropic API script: replaces `extractions.py` for new TRID exports. Uses Sonnet 4.5, strict JSON schema, moderate stance on ML gaps. Supports resume. |
| `build_excel.py` | Assembles the Excel workbook from `records_clean.json` + `extractions.py` (or the JSON produced by `extract_via_api.py`). |
| `trid_review_matrix.xlsx` | The final deliverable: Research Matrix + Summary + Notes sheets. |

## Rerun on a new TRID export

```bash
# 1. Put your new exports into a known folder
# 2. Parse
python parse_trid.py                  # writes records.json
# Manually drop empty records -> records_clean.json  (or reuse the script logic in the notes)

# 3. Extract research content via API
export ANTHROPIC_API_KEY=sk-ant-...
python extract_via_api.py             # writes extractions.json

# 4. If using extract_via_api.py, edit build_excel.py line:
#     from extractions import EXTRACTIONS
# to instead load JSON:
#     EXTRACTIONS = json.load(open("extractions.json"))

# 5. Build Excel
python build_excel.py                 # writes trid_review_matrix.xlsx
```

Rough cost: ~$2–5 for 120 abstracts at Sonnet 4.5 prices.

## Column schema

**Metadata** (dark blue): ID, Title, Authors, Year, Venue
**Geography** (medium blue): Country, City, Geography_Bucket
**Method** (green): Simulation_Type, Software, Modeling_Method, Freight_Segment, Data_Source
**Policy** (orange): Policy_Lever, Key_Finding, Stated_Limitations
**Gap** (red-orange): ML_AI_Gap

## What I noticed across the 119 papers (for your gap analysis)

**Geography skew is real and quantifiable.** The largest clusters are: Europe (EU) ~40, Asia ~25, USA ~20, Europe (other) ~10, Oceania + LATAM + Canada + Africa making up the rest. Your intuition that most are outside the USA is correct — about 83% are non-USA.

**Method concentration.** Agent-based simulation (MATSim, SimMobility Freight, MASS-GT, custom ABM) dominates post-2016. Pre-2015 is heavier on discrete-event and system dynamics. Monte Carlo is resurgent for e-commerce / electrification TCO analyses. Cellular automata appears only in the Polish Szczecin group. Traffic microsimulation (VISSIM, TransModeler) is the choice for corridor-level or signal-timing work.

**Policy silos.** Almost every paper targets a single lever: UCC, LEZ/zero-emission zone, cargo bikes, electrification, off-hour/PierPASS, platooning, crowdshipping, loading bays, consolidation, time windows, parking/pricing. Cross-policy interaction is rare — that is itself a gap worth claiming in a dissertation.

**Biggest repeated ML/AI openings:**
1. **Surrogate models** to replace expensive simulation runs — explicit in the TRB project on network-level truck reliability (#117, hundreds of runs needed).
2. **ML calibration** replacing hand-tuned PSO/GA (#12 SimMobility screenline calibration, #83 parallel PSO).
3. **Demand synthesis from passive data** (mobile LBS, GPS, OSM) instead of expensive establishment surveys — cited as the limiting constraint in #8, #47, #55, #62, #81.
4. **Choice models** — most of these papers use MNL/nested logit; graph neural networks or transformer-based choice models could capture network effects (supplier-receiver relationships, tour chaining).
5. **Reinforcement learning for multi-stakeholder policy design** — only #45, #57 explicitly use ADP. RL for joint pricing + cooperation + access restrictions is wide open.
6. **Counterfactual / causal ML** for policy-combination attribution — multiple papers test single policies but warn their effects may not compose (#35, #43, #89).

These are the threads to pull on when you write your literature-review chapter.
