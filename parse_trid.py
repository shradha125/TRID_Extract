"""Parse TRID-exported .docx files into structured bibliographic records."""
import re
import json
import subprocess
from pathlib import Path


def extract_text(docx_path: str) -> str:
    """Use the extract-text CLI (pandoc-backed) to get markdown."""
    result = subprocess.run(
        ["extract-text", docx_path], capture_output=True, text=True, check=True
    )
    return result.stdout


AUTHOR_LINE_RE = re.compile(r"^[A-ZÀ-Ý][A-ZÀ-Ý\s,;.\-'’`]+$")  # ALL-CAPS author-style line

def is_likely_author_line(s: str) -> bool:
    """Heuristic: lines like 'VISSER, JGSN; MAAT, C' or 'Smith, John; Doe, Jane' are authors, not titles."""
    content = s.strip("* ").strip()
    if not content:
        return False
    # Has comma+space pattern typical of 'Lastname, Firstname'
    if re.search(r"[A-ZÀ-Ý][a-zà-ÿ]+,\s+[A-ZÀ-Ý]", content):
        return True
    # ALL-CAPS short line with commas/semicolons = author list
    if AUTHOR_LINE_RE.match(content) and ("," in content or ";" in content):
        return True
    return False


def split_records(text: str) -> list[str]:
    """Each TRID record is anchored by an '**Abstract**.' marker; we walk back to find the title."""
    abstract_positions = [m.start() for m in re.finditer(r"\*\*Abstract\*\*\.", text)]
    records = []
    for i, pos in enumerate(abstract_positions):
        prefix = text[:pos]
        lines = prefix.rstrip().split("\n")
        title = None
        j = len(lines) - 1
        # Walk back past blanks, tab-indented author lines, and flush-left author lines
        while j >= 0:
            ln = lines[j].rstrip()
            if not ln.strip():
                j -= 1
                continue
            if not (ln.startswith("**") and ln.endswith("**")):
                j -= 1
                continue
            # Skip if it looks like an author line (tab-indented OR matches author-name pattern)
            if "\t" in ln or is_likely_author_line(ln):
                j -= 1
                continue
            # Otherwise this is the title
            title = ln.strip("* ").strip()
            break
        # End of this record is the start of the next abstract's title OR end of text
        if i + 1 < len(abstract_positions):
            end = abstract_positions[i + 1]
            next_prefix = text[:end]
            next_lines = next_prefix.rstrip().split("\n")
            k = len(next_lines) - 1
            while k >= 0:
                ln = next_lines[k].rstrip()
                if (ln.startswith("**") and ln.endswith("**") and "\t" not in ln
                        and not is_likely_author_line(ln) and ln.strip("* ").strip()):
                    end = next_prefix.rfind(ln)
                    break
                k -= 1
            record_text = text[:end]
        else:
            record_text = text
        # Trim to just this record: from last title before pos to end
        if title:
            # Find last occurrence of title in prefix
            title_pos = prefix.rfind(f"**{title}**")
            if title_pos >= 0:
                records.append(record_text[title_pos:])
    return records


def parse_field(label: str, text: str) -> str | None:
    """Extract a '**Label:** value' style field (value can span to next ** or blank line)."""
    pattern = rf"\*\*{re.escape(label)}:\*\*\s*(.*?)(?=\n\n|\*\*[A-Z])"
    m = re.search(pattern, text, re.DOTALL)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip(" *\n\t")
    return None


def parse_record(text: str) -> dict:
    """Extract structured fields from a single record's markdown."""
    lines = text.split("\n")
    title = lines[0].strip("* ").strip() if lines else ""

    # Authors: the bold line(s) BETWEEN title and **Abstract**. For projects there is no author line.
    authors = ""
    abs_idx = text.find("**Abstract**")
    if abs_idx > 0:
        between = text[len(lines[0]):abs_idx]
        for ln in between.split("\n"):
            s = ln.strip()
            if not s or not s.startswith("**"):
                continue
            if s.startswith("**Abstract") or s.startswith("**Record"):
                continue
            content = re.sub(r"\*+", "", s).strip()
            # Author lines typically contain commas/semicolons and are not sentence-like
            if ("," in content or ";" in content) and len(content) < 300 and "." not in content[:-1]:
                authors = re.sub(r"\s+", " ", content)
                break

    # Abstract
    abs_match = re.search(r"\*\*Abstract\*\*\.\s*(.*?)(?=\n\*\*Record Type|\n\*\*Record URL|\Z)", text, re.DOTALL)
    abstract = re.sub(r"\s+", " ", abs_match.group(1)).strip() if abs_match else ""

    # Publication date
    pub_date = parse_field("Publication Date", text) or ""
    year_match = re.search(r"(19|20)\d{2}", pub_date)
    year = year_match.group(0) if year_match else ""
    if not year:
        start = parse_field("Start Date", text) or ""
        ym = re.search(r"(19|20)\d{2}", start)
        year = ym.group(0) if ym else ""

    return {
        "title": title,
        "authors": authors,
        "abstract": abstract,
        "year": year,
        "record_type": parse_field("Record Type", text) or "",
        "record_url": parse_field("Record URL", text) or "",
        "serial": parse_field("Serial", text) or "",
        "subject_areas": parse_field("Subject Areas", text) or "",
        "keywords": parse_field("Keywords", text) or "",
        "sponsor": parse_field("Sponsor Organizations", text) or "",
        "performing_org": parse_field("Performing Organizations", text) or "",
        "pi": parse_field("Principal Investigators", text) or "",
    }


def main():
    all_records = []
    for fp in [
        "/mnt/user-data/uploads/simulation_articles_and_papers_cut.docx",
        "/mnt/user-data/uploads/simulation_projects.docx",
    ]:
        text = extract_text(fp)
        records = split_records(text)
        print(f"{Path(fp).name}: {len(records)} records")
        for r in records:
            all_records.append(parse_record(r))

    with open("/home/claude/records.json", "w") as f:
        json.dump(all_records, f, indent=2)
    print(f"Total: {len(all_records)} records saved")


if __name__ == "__main__":
    main()
