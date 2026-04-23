"""
Reusable API-based TRID research-content extractor.

Run locally with your Anthropic API key set as ANTHROPIC_API_KEY.
Feeds each parsed record's abstract to Claude and asks for a strict JSON
object matching the 12-field extraction schema, then writes extractions.json
that build_excel.py can consume.

Usage:
    export ANTHROPIC_API_KEY=sk-ant-...
    python parse_trid.py                 # produces records.json
    # clean records.json -> records_clean.json as in README
    python extract_via_api.py            # produces extractions.json
    python build_excel.py                # produces trid_review_matrix.xlsx
"""
import json
import os
import sys
import time
from pathlib import Path

try:
    from anthropic import Anthropic
except ImportError:
    print("pip install anthropic", file=sys.stderr)
    sys.exit(1)

MODEL = "claude-sonnet-4-5"  # fast enough for ~120 records
INPUT_FILE = "records_clean.json"
OUTPUT_FILE = "extractions.json"

SYSTEM_PROMPT = """You are a transportation-engineering research assistant extracting structured \
fields from micro-simulation freight-transport paper abstracts. Return STRICT JSON with exactly \
these keys and no others: Country, City, Simulation_Type, Software, Modeling_Method, \
Policy_Lever, Freight_Segment, Data_Source, Key_Finding, Stated_Limitations, ML_AI_Gap, \
Geography_Bucket.

Rules:
- Geography_Bucket MUST be one of: USA | North America (Canada) | Europe (EU) | Europe (other) | \
Asia | Oceania | LATAM | Africa | Middle East
- Use 'Not stated' when a field cannot be inferred from the abstract. Never invent facts.
- Key_Finding = one sentence stating the quantitative result.
- Stated_Limitations = authors' own acknowledged gaps, not your critique.
- ML_AI_Gap uses a MODERATE stance: propose one concrete ML/AI extension that addresses a \
stated limitation or obvious methodological weakness (e.g., neural surrogate for expensive sim \
runs, RL for adaptive policy, GNN for network effects, CV for automated data collection, \
Bayesian optimization for calibration). Do not claim the authors failed to use ML if they did.
- Simulation_Type examples: "Agent-based", "Monte Carlo", "Discrete-event", "Traffic microsimulation", \
"System dynamics", "Cellular automata", "Hybrid ABM/DES", etc.
- Return ONLY the JSON object. No prose, no code fences."""

USER_TEMPLATE = """TITLE: {title}
AUTHORS: {authors}
YEAR: {year}
KEYWORDS: {keywords}

ABSTRACT:
{abstract}"""

def extract_one(client: Anthropic, rec: dict) -> dict:
    msg = client.messages.create(
        model=MODEL,
        max_tokens=1200,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": USER_TEMPLATE.format(
                title=rec["title"],
                authors=rec["authors"],
                year=rec["year"],
                keywords=rec.get("keywords", ""),
                abstract=rec["abstract"],
            ),
        }],
    )
    text = msg.content[0].text.strip()
    # Strip accidental code fences if the model slips up
    if text.startswith("```"):
        text = text.strip("`").split("\n", 1)[1].rsplit("\n", 1)[0]
        if text.startswith("json"):
            text = text[4:].lstrip()
    return json.loads(text)

def main():
    records = json.load(open(INPUT_FILE))
    # Resume support: if output exists, pick up where we left off
    if Path(OUTPUT_FILE).exists():
        done = json.load(open(OUTPUT_FILE))
    else:
        done = []
    start = len(done)
    print(f"Resuming at record {start}/{len(records)}")

    client = Anthropic()
    for i in range(start, len(records)):
        rec = records[i]
        if not rec["abstract"].strip():
            done.append({"error": "empty abstract"})
            continue
        for attempt in range(3):
            try:
                result = extract_one(client, rec)
                done.append(result)
                break
            except Exception as e:
                print(f"  record {i} attempt {attempt+1} failed: {e}", file=sys.stderr)
                time.sleep(2 ** attempt)
        else:
            done.append({"error": "failed after retries"})
        # Write incrementally so resume works
        json.dump(done, open(OUTPUT_FILE, "w"), indent=2)
        print(f"[{i+1}/{len(records)}] {rec['title'][:70]}")

    print(f"Done. {len(done)} extractions in {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
