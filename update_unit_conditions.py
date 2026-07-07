"""
Weekly Unit Conditions Updater
==============================
Pulls the latest "Oaks at Creekside Weekly Call Meeting Minutes" email from Gmail,
extracts all unit-specific notes, and appends them to the Unit Conditions tab.

Usage (you just tell Claude: "Update unit conditions from this week's meeting"):
    python update_unit_conditions.py

Or for a specific date:
    python update_unit_conditions.py --date "June 24, 2026"

Or dry-run (see what would be written without touching the sheet):
    python update_unit_conditions.py --dry-run
"""

import os, sys, re, json, time, argparse
from datetime import datetime
from html.parser import HTMLParser

sys.path.insert(0, os.path.dirname(__file__))
from creekside_automation import get_gspread_client, SPREADSHEET_ID

UNIT_CONDITIONS_SHEET = "Unit Conditions"
EMAIL_SUBJECT_PREFIX  = "Oaks at Creekside Weekly Call Meeting Minutes"

# ── HTML → plain text ───────────────────────────────────────────────────────

class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.result = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag in ("li",):    self.result.append("\n• ")
        if tag in ("p","h2","h3","h4","br"): self.result.append("\n")

    def handle_data(self, data):
        self.result.append(data)

    def get_text(self):
        return re.sub(r"\n{3,}", "\n\n", "".join(self.result)).strip()


def html_to_text(html):
    s = _HTMLStripper()
    s.feed(html)
    return s.get_text()


# ── Gmail fetch ─────────────────────────────────────────────────────────────

def fetch_latest_meeting_email(target_date_str=None):
    """Return (subject, date_str, plaintext_body) of the most recent meeting email."""
    try:
        import gspread  # just to confirm environment
        # Use Gmail MCP via subprocess isn't possible here — caller must inject body
        # This function is a placeholder when running stand-alone.
        # When Claude runs this, it fetches the email and passes the body directly.
        raise NotImplementedError("Direct Gmail fetch requires Claude session — pass --body flag or let Claude call this.")
    except NotImplementedError:
        raise


# ── Unit notes parser ────────────────────────────────────────────────────────

# Sections in the email that contain unit-specific notes
UNIT_SECTION_HEADERS = [
    "unit conditions",
    "unit make-ready",
    "make-ready and maintenance",
    "recent move-outs",
    "unit turn",
]

def extract_unit_notes_from_text(text, meeting_date_label):
    """
    Parse plain-text meeting notes. Returns dict: unit_number -> note_text.
    Looks for patterns like:
      - "Unit 202 (Skipped):" followed by bullet points
      - "Unit 301 (Roach Issue):" followed by bullet points
    """
    units = {}

    # Find all "Unit NNN" mentions with following bullet content
    # Pattern: Unit ### (optional label): followed by indented/bulleted lines
    unit_block_re = re.compile(
        r'[•\-*]?\s*\*?Unit\s+(\d{3,4})\*?(?:\s*[\(\[]?[^\n\)]{0,40}[\)\]]?)?\s*[:\-]?\s*\n'
        r'((?:(?:[ \t]*[•\-*•]?[ \t]+[^\n]+\n?)+))',
        re.IGNORECASE | re.MULTILINE
    )

    for m in unit_block_re.finditer(text):
        unit_num = m.group(1)
        raw_block = m.group(2)
        # Clean up bullets and indentation
        lines = []
        for line in raw_block.splitlines():
            line = re.sub(r'^[\s•\-*•]+', '', line).strip()
            if line:
                lines.append(line)
        if lines:
            note = meeting_date_label + "\n" + "\n".join(lines)
            # Merge if unit appears multiple times
            if unit_num in units:
                units[unit_num] += "\n" + "\n".join(lines)
            else:
                units[unit_num] = note

    return units


def extract_unit_notes_structured(text, meeting_date_label):
    """
    Secondary parser: look for bold/header unit references in structured sections.
    Handles format: '• Unit 202 (Skipped):' then sub-bullets indented.
    """
    units = {}

    # Split into lines and scan
    lines = text.splitlines()
    current_unit = None
    current_lines = []
    in_unit_section = False

    for i, line in enumerate(lines):
        stripped = line.strip()

        # Detect section headers that contain unit conditions
        lower = stripped.lower()
        if any(h in lower for h in UNIT_SECTION_HEADERS):
            in_unit_section = True

        # Detect a unit heading like "Unit 202", "Unit 1416:", "* Unit 301 (Roach Issue)"
        unit_match = re.match(
            r'^[•\-*•\s]*\*?Unit\s+(\d{3,4})\*?(?:\s*[\(\[][^\)\]]*[\)\]])?\s*[:\-]?\s*$',
            stripped, re.IGNORECASE
        )
        if unit_match:
            # Save previous unit
            if current_unit and current_lines:
                note = meeting_date_label + "\n" + "\n".join(current_lines)
                if current_unit in units:
                    units[current_unit] += "\n" + "\n".join(current_lines)
                else:
                    units[current_unit] = note
            current_unit = unit_match.group(1)
            current_lines = []
            continue

        # Collect sub-bullets under current unit
        if current_unit:
            # Stop if we hit a new section header or empty line followed by non-bullet
            if stripped == "" and i + 1 < len(lines):
                next_stripped = lines[i + 1].strip()
                # If next line looks like another unit or a section header, stop
                if re.match(r'^[•\-*•\s]*\*?Unit\s+\d', next_stripped, re.IGNORECASE):
                    continue
                if any(h in next_stripped.lower() for h in ["action items","key performance","discussion","vendor","delinquency"]):
                    # Save and stop collecting for this unit
                    if current_lines:
                        note = meeting_date_label + "\n" + "\n".join(current_lines)
                        if current_unit in units:
                            units[current_unit] += "\n" + "\n".join(current_lines)
                        else:
                            units[current_unit] = note
                        current_lines = []
                    current_unit = None
                    continue
            if stripped:
                clean = re.sub(r'^[•\-*•]+\s*', '', stripped)
                if clean:
                    current_lines.append(clean)

    # Save last unit
    if current_unit and current_lines:
        note = meeting_date_label + "\n" + "\n".join(current_lines)
        if current_unit in units:
            units[current_unit] += "\n" + "\n".join(current_lines)
        else:
            units[current_unit] = note

    return units


# Lines matching these patterns are scheduling/tenancy matters — not unit conditions
_NON_CONDITION_PATTERNS = re.compile(
    r'\b('
    r'make.?ready\s+(completed?|done|finished|started|not\s+started|in\s+progress)'
    r'|completed?\s+make.?ready'
    r'|just\s+needs?\s+cleaning'   # keep "needs cleaning" but strip "just"? No — cleaning IS a condition
    r'|switching\s+to\s+unit'
    r'|skip(ped)?\s*[-—]'
    r'|eviction'
    r'|court\s+date'
    r'|family\s+emergency'
    r'|no\s+longer\s+working'
    r'|will\s+be\s+out'
    r'|no\s+payments?\s+since'
    r'|resident\s+said\s+will\s+be\s+out'
    r')\b',
    re.IGNORECASE,
)

# A line is a physical condition if it mentions something tangible about the unit
_CONDITION_KEYWORDS = re.compile(
    r'\b('
    r'leak|damage|mold|pest|roach|bug|infestation|flood|water|stain'
    r'|crack|broken|broken|hole|missing|replaced?|repair|fix|issue|problem'
    r'|paint|floor|carpet|plank|tile|ceiling|wall|door|window|hvac|a/?c|heat'
    r'|appliance|stove|fridge|refrigerator|dishwasher|washer|dryer|disposal'
    r'|cabinet|counter|sink|tub|toilet|shower|plumbing|electrical|outlet'
    r'|cleaning|clean|trash|debris|odor|smell|smoke|biohazard'
    r'|roof|exterior|balcony|patio|foundation|structural'
    r')\b',
    re.IGNORECASE,
)


def filter_physical_conditions(note_text, date_label):
    """
    Given a raw note block (starting with date_label), return only the lines
    that describe physical unit conditions. Returns None if nothing remains.
    """
    lines = note_text.splitlines()
    kept = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped == date_label:
            continue
        has_condition = bool(_CONDITION_KEYWORDS.search(stripped))
        has_non_condition = bool(_NON_CONDITION_PATTERNS.search(stripped))
        # Physical condition keyword always wins — keep the line
        if has_condition:
            kept.append(stripped)
        # Drop lines that are purely scheduling/tenancy with no condition content
        elif has_non_condition:
            continue
        # Keep other descriptive lines that aren't explicitly administrative
        elif not re.search(
            r'\b(completed?|started|not\s+started|in\s+progress|switching|skip(ped)?|eviction|court|emergency|no\s+longer|will\s+be\s+out)\b',
            stripped, re.IGNORECASE
        ):
            if len(stripped) > 4:
                kept.append(stripped)

    if not kept:
        return None
    return date_label + "\n" + "\n".join(kept)


def parse_meeting_notes(email_body, meeting_date_label):
    """
    Master parser — runs both parsers, merges results, then filters to
    physical condition details only (strips evictions, skips, scheduling).
    Returns dict: unit_num -> formatted note string (starting with date).
    """
    text = email_body if not email_body.strip().startswith("<") else html_to_text(email_body)

    notes_a = extract_unit_notes_from_text(text, meeting_date_label)
    notes_b = extract_unit_notes_structured(text, meeting_date_label)

    # Merge: prefer whichever has more content per unit
    merged = {}
    all_units = set(notes_a) | set(notes_b)
    for u in all_units:
        a = notes_a.get(u, "")
        b = notes_b.get(u, "")
        merged[u] = a if len(a) >= len(b) else b

    # Filter to physical conditions only
    filtered = {}
    for u, note in merged.items():
        clean = filter_physical_conditions(note, meeting_date_label)
        if clean:
            filtered[u] = clean
        else:
            print(f"  [skip] Unit {u}: no physical condition details — not written to sheet")

    return filtered


# ── Sheet updater ────────────────────────────────────────────────────────────

def update_sheet(unit_notes, meeting_date_label, dry_run=False):
    """
    Append new notes to Condition / Issue column (col D) for each unit.
    Creates a new row if unit not found in sheet.
    Returns summary dict.
    """
    gc = get_gspread_client()
    wb = gc.open_by_key(SPREADSHEET_ID)
    ws = wb.worksheet(UNIT_CONDITIONS_SHEET)
    vals = ws.get_all_values()

    # Build unit -> row index (1-based) map
    unit_to_row = {}
    for i, row in enumerate(vals[1:], start=2):
        u = row[0].strip() if row else ""
        if u:
            unit_to_row[u] = i

    results = {"updated": [], "created": [], "skipped": [], "no_match": []}
    batch_updates = []
    new_rows = []

    for unit, new_note in unit_notes.items():
        date_tag = meeting_date_label  # e.g. "6/24/26"

        if unit in unit_to_row:
            row_idx = unit_to_row[unit]
            existing = vals[row_idx - 1][3] if len(vals[row_idx - 1]) > 3 else ""

            if date_tag in existing:
                results["skipped"].append(unit)
                continue

            separator = "\n\n" if existing.strip() else ""
            updated_val = existing.strip() + separator + new_note
            batch_updates.append({
                "range":  f"D{row_idx}",
                "values": [[updated_val]]
            })
            results["updated"].append(unit)
        else:
            # New unit — append row
            new_rows.append([unit, "", "", new_note, "", ""])
            results["created"].append(unit)

    print(f"\n[Unit Conditions — {meeting_date_label}]")
    print(f"  Units to update : {len(results['updated'])} — {results['updated']}")
    print(f"  Units to create : {len(results['created'])} — {results['created']}")
    print(f"  Already current : {len(results['skipped'])} — {results['skipped']}")

    if dry_run:
        print("  (Dry run — no changes written)")
        for unit, note in unit_notes.items():
            print(f"\n  --- Unit {unit} ---\n{note}")
        return results

    if batch_updates:
        ws.batch_update(batch_updates, value_input_option="USER_ENTERED")
        print(f"  ✅ Updated {len(batch_updates)} rows")
        time.sleep(1)

    if new_rows:
        ws.append_rows(new_rows, value_input_option="USER_ENTERED")
        print(f"  ✅ Created {len(new_rows)} new rows")

    return results


# ── Main ─────────────────────────────────────────────────────────────────────

def run_from_email_body(email_body, meeting_date_label, dry_run=False):
    """
    Entry point used by Claude: pass in email body + date label.
    Returns summary.
    """
    print(f"Parsing meeting notes for {meeting_date_label}…")
    unit_notes = parse_meeting_notes(email_body, meeting_date_label)
    print(f"Found {len(unit_notes)} units with notes: {sorted(unit_notes.keys())}")

    if not unit_notes:
        print("No unit notes found in email — check parsing.")
        return {}

    result = update_sheet(unit_notes, meeting_date_label, dry_run=dry_run)

    if not dry_run and (result.get("updated") or result.get("created")):
        try:
            from build_dashboard import main as build_dashboard
            print("\nRebuilding dashboard with updated notes…")
            build_dashboard()
        except Exception as e:
            print(f"[WARN] Dashboard not rebuilt: {e}")

    return result


def main():
    parser = argparse.ArgumentParser(description="Update Unit Conditions from weekly meeting email")
    parser.add_argument("--date",    help="Meeting date string, e.g. 'June 24, 2026'")
    parser.add_argument("--body",    help="Path to plain-text file with email body")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    if not args.body:
        print("Pass --body path/to/email.txt with the email body, or let Claude run this automatically.")
        sys.exit(1)

    with open(args.body) as f:
        body = f.read()

    date_label = args.date or datetime.now().strftime("%-m/%-d/%y")
    run_from_email_body(body, date_label, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
