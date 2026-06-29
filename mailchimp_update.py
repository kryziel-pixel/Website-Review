"""
Mailchimp Template Updater
--------------------------
Reads a PDF or Word (.docx) monthly update file and updates a Mailchimp
template with the extracted content, preserving all formatting/layout.

Usage:
    # Inspect what the script finds (dry run):
    python3 mailchimp_update.py --file "May 26 Update.pdf" --template-id 10105010 --dry-run

    # Update the template directly:
    python3 mailchimp_update.py --file "May 26 Update.pdf" --template-id 10105010

    # Create a new campaign from a template, then fill it:
    python3 mailchimp_update.py --file "May 26 Update.pdf" --template-id 10105010 \
        --duplicate --list-id <audience-id> --subject "Ranch at Sienna – June Update"

    # Update an already-created campaign by its ID:
    python3 mailchimp_update.py --file "May 26 Update.pdf" --campaign-id <uuid>

Requirements:
    pip install requests beautifulsoup4 pdfplumber python-docx python-dotenv lxml

Environment (.env file in same folder):
    MAILCHIMP_API_KEY=your-api-key-us3
"""

import argparse
import os
import re
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

API_KEY = os.getenv("MAILCHIMP_API_KEY", "")
SERVER  = os.getenv("MAILCHIMP_SERVER", "")

if not SERVER and "-" in API_KEY:
    SERVER = API_KEY.split("-")[-1]

BASE_URL = f"https://{SERVER}.api.mailchimp.com/3.0"


def api(method: str, path: str, **kwargs):
    url = BASE_URL + path
    resp = requests.request(method, url, auth=("anystring", API_KEY), **kwargs)
    if not resp.ok:
        print(f"[ERROR] {method} {path} → {resp.status_code}: {resp.text[:500]}")
        sys.exit(1)
    return resp.json() if resp.text else {}


# ---------------------------------------------------------------------------
# PDF / DOCX extraction
# ---------------------------------------------------------------------------

def extract_text_from_pdf(path: str) -> str:
    try:
        import pdfplumber
    except ImportError:
        sys.exit("pdfplumber not installed. Run: pip3 install pdfplumber")

    text = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text.append(page.extract_text() or "")
    return "\n".join(text)


def extract_text_from_docx(path: str) -> str:
    try:
        from docx import Document
    except ImportError:
        sys.exit("python-docx not installed. Run: pip3 install python-docx")

    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)


def extract_text(path: str) -> str:
    ext = Path(path).suffix.lower()
    if ext == ".pdf":
        return extract_text_from_pdf(path)
    elif ext in (".docx", ".doc"):
        return extract_text_from_docx(path)
    else:
        sys.exit(f"Unsupported file type: {ext}. Use .pdf or .docx")


# ---------------------------------------------------------------------------
# Section parsing
# ---------------------------------------------------------------------------

SECTION_PATTERNS = {
    "executive_summary": r"Executive Summary\s*(.*?)(?=\nOperations\b|$)",
    "operations":        r"\nOperations\s*\n(.*?)(?=\nRevenue Initiatives|$)",
    "revenue":           r"Revenue Initiatives[:\s]*(.*?)(?=\nFinancials\b|$)",
    "financials":        r"\nFinancials\s*\n(.*?)$",
}

SECTION_HEADINGS = {
    "executive_summary": "Executive Summary",
    "operations":        "Operations",
    "revenue":           "Revenue Initiatives:",
    "financials":        "Financials",
}


def parse_sections(text: str) -> dict:
    sections = {}
    for key, pattern in SECTION_PATTERNS.items():
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if match:
            content = match.group(1).strip()
            content = re.sub(r"\n{3,}", "\n\n", content)
            sections[key] = content
    return sections


def text_to_paragraphs(text: str) -> list:
    """Split text into paragraph strings."""
    return [p.strip() for p in text.split("\n\n") if p.strip()]


# ---------------------------------------------------------------------------
# Mailchimp API helpers
# ---------------------------------------------------------------------------

def get_template_html(template_id: int) -> str:
    """Fetch the rendered HTML of a template via its default content."""
    print(f"Fetching template {template_id}...")
    # Try default-content first
    data = api("GET", f"/templates/{template_id}/default-content")
    html = data.get("html", "")
    if html:
        return html
    # Fall back to the template record itself
    tpl = api("GET", f"/templates/{template_id}")
    return tpl.get("html", "")


def update_template_html(template_id: int, html: str):
    print(f"Saving updated template {template_id}...")
    api("PATCH", f"/templates/{template_id}", json={"html": html})
    print(f"  Done: https://{SERVER}.admin.mailchimp.com/templates/edit?id={template_id}")


def get_campaign_content_html(campaign_id: str) -> str:
    print(f"Fetching campaign {campaign_id} content...")
    data = api("GET", f"/campaigns/{campaign_id}/content")
    return data.get("html", "")


def update_campaign_html(campaign_id: str, html: str):
    print(f"Saving campaign {campaign_id}...")
    api("PUT", f"/campaigns/{campaign_id}/content", json={"html": html})
    print(f"  Done.")


def create_campaign_from_template(template_id: int, subject: str, list_id: str) -> str:
    print(f"Creating campaign from template {template_id}...")
    payload = {
        "type": "regular",
        "recipients": {"list_id": list_id},
        "settings": {
            "subject_line": subject,
            "title":        subject,
            "template_id":  template_id,
        },
    }
    resp = api("POST", "/campaigns", json=payload)
    cid = resp["id"]
    print(f"  Campaign created: {cid}")
    return cid


# ---------------------------------------------------------------------------
# HTML content replacement
# ---------------------------------------------------------------------------

# Rough text content of each heading as it appears in the email HTML
HEADING_SIGNATURES = {
    "executive_summary": ["executive summary"],
    "operations":        ["operations"],
    "revenue":           ["revenue initiatives"],
    "financials":        ["financials"],
}

# Order of sections so we know where each one ends
SECTION_ORDER = ["executive_summary", "operations", "revenue", "financials"]


def get_text(tag) -> str:
    return tag.get_text(" ", strip=True).lower()


def find_section_heading_tags(soup: BeautifulSoup) -> dict:
    """
    Find the HTML tag that acts as the heading for each section.
    Returns {section_key: tag}.
    """
    heading_tags = {}
    candidates = soup.find_all(["h1", "h2", "h3", "h4", "p", "div", "td", "span"])
    for tag in candidates:
        text = get_text(tag)
        for key, sigs in HEADING_SIGNATURES.items():
            if key not in heading_tags:
                for sig in sigs:
                    if text == sig or text.rstrip(":") == sig:
                        heading_tags[key] = tag
                        break
    return heading_tags


def find_content_container(heading_tag) -> Tag | None:
    """
    Walk up from a heading tag to find the table row or section wrapper
    that contains the heading, then return the NEXT sibling container
    that holds the body text.
    """
    # Try next sibling in parent
    parent = heading_tag.parent
    siblings = list(parent.children)
    idx = siblings.index(heading_tag) if heading_tag in siblings else -1

    # Walk up until we find a <tr> or <table> parent with useful siblings
    node = heading_tag
    for _ in range(6):
        sibling = node.find_next_sibling()
        if sibling and sibling.name in ("p", "div", "td", "table", "tr"):
            return sibling
        node = node.parent
        if node is None:
            break
    return None


def replace_text_preserving_style(container: Tag, new_paragraphs: list):
    """
    Replace text content inside a container tag while keeping the first
    existing <p> or <td> as a style donor (font, color, size attributes).
    """
    # Find style donor — first <p> or text-bearing tag inside container
    donor = container.find("p") or container.find("td") or container.find("div")

    # Clear container
    container.clear()

    soup_frag = BeautifulSoup("", "lxml")

    for i, para_text in enumerate(new_paragraphs):
        if donor:
            new_tag = BeautifulSoup(str(donor), "lxml").find(donor.name)
            if new_tag:
                new_tag.clear()
                new_tag.string = para_text
                container.append(new_tag.__copy__() if hasattr(new_tag, "__copy__") else BeautifulSoup(str(new_tag), "lxml").find(donor.name))
                continue
        # Fallback: plain <p>
        p = BeautifulSoup(f"<p>{para_text}</p>", "lxml").find("p")
        container.append(p)


def inject_sections_into_html(html: str, sections: dict, dry_run: bool = False) -> str:
    """
    Strategy:
    1. Find each section heading in the HTML.
    2. Locate the nearest content block after it.
    3. Replace the text in that block with new content.
    """
    soup = BeautifulSoup(html, "lxml")
    heading_tags = find_section_heading_tags(soup)

    if not heading_tags:
        print("[WARN] Could not find any section headings in the template HTML.")
        print("       The template may use image-based or highly custom layout.")
        print("       Falling back to full text search replacement...")
        return replace_by_text_search(html, sections, dry_run)

    print(f"\nFound headings: {list(heading_tags.keys())}")

    for key in SECTION_ORDER:
        if key not in sections or key not in heading_tags:
            continue

        heading_tag = heading_tags[key]
        new_paragraphs = text_to_paragraphs(sections[key])

        # Find next content block after this heading
        container = find_content_container(heading_tag)

        if dry_run:
            print(f"\n[{key}] Heading found: {get_text(heading_tag)!r}")
            print(f"  Container: {container.name if container else 'NOT FOUND'}")
            print(f"  New content preview: {new_paragraphs[0][:80] if new_paragraphs else '(empty)'}...")
            continue

        if container:
            replace_text_preserving_style(container, new_paragraphs)
            print(f"  [{key}] Updated.")
        else:
            print(f"  [{key}] Could not find content container after heading — skipping.")

    return str(soup)


def replace_by_text_search(html: str, sections: dict, dry_run: bool = False) -> str:
    """
    Fallback: search for known phrases from the old email and replace paragraphs.
    Uses the first sentence of each section as a search anchor.
    """
    # Known first sentences from the previous month's email (update these if template changes)
    ANCHORS = {
        "executive_summary": "Occupancy closed the month",
        "operations":        "Physical occupancy finished the month",
        "revenue":           "We continue to evaluate operational strategies",
        "financials":        "Total Income for",
    }

    soup = BeautifulSoup(html, "lxml")

    for key, anchor in ANCHORS.items():
        if key not in sections:
            continue
        # Find the tag containing this anchor text
        tag = soup.find(lambda t: t.string and anchor.lower() in t.get_text().lower()
                        and t.name in ("p", "td", "div", "span"))
        if tag:
            new_paragraphs = text_to_paragraphs(sections[key])
            if dry_run:
                print(f"[{key}] Found anchor in <{tag.name}>: {tag.get_text()[:60]}...")
                print(f"  → Will replace with: {new_paragraphs[0][:80]}...")
            else:
                parent = tag.parent
                # Replace all sibling <p> tags after anchor until next heading
                tag.string = new_paragraphs[0]
                next_sib = tag.find_next_sibling()
                for extra in new_paragraphs[1:]:
                    new_p = BeautifulSoup(f"<p>{extra}</p>", "lxml").find("p")
                    tag.insert_after(new_p)
                    tag = new_p
                print(f"  [{key}] Updated via text search.")
        else:
            print(f"  [{key}] Anchor not found in HTML — skipping.")

    return str(soup)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Update a Mailchimp template from a PDF/Word file.")
    parser.add_argument("--file",          required=True, help="Path to PDF or .docx update file")
    parser.add_argument("--template-id",   type=int,      help="Mailchimp template ID")
    parser.add_argument("--campaign-id",                  help="Mailchimp campaign UUID")
    parser.add_argument("--duplicate",     action="store_true",
                        help="Create a new campaign from the template first")
    parser.add_argument("--list-id",       help="Audience/list ID (required with --duplicate)")
    parser.add_argument("--subject",       help="Email subject line (required with --duplicate)")
    parser.add_argument("--dry-run",       action="store_true",
                        help="Parse and detect sections but don't write anything to Mailchimp")
    args = parser.parse_args()

    if not API_KEY:
        sys.exit("Set MAILCHIMP_API_KEY in your .env file.")

    # --- Extract and parse content ---
    print(f"Reading {args.file}...")
    raw_text  = extract_text(args.file)
    sections  = parse_sections(raw_text)

    print(f"\nDetected {len(sections)} sections: {list(sections.keys())}")
    for k, v in sections.items():
        print(f"  [{k}] {v[:100].replace(chr(10), ' ')}...")

    if args.dry_run:
        print("\n--- DRY RUN — no changes will be written ---")

    # --- Determine HTML source ---
    if args.campaign_id:
        html = get_campaign_content_html(args.campaign_id)
        updated = inject_sections_into_html(html, sections, dry_run=args.dry_run)
        if not args.dry_run:
            update_campaign_html(args.campaign_id, updated)

    elif args.template_id:
        if args.duplicate:
            if not args.list_id or not args.subject:
                sys.exit("--duplicate requires --list-id and --subject")
            cid = create_campaign_from_template(args.template_id, args.subject, args.list_id)
            html = get_campaign_content_html(cid)
            updated = inject_sections_into_html(html, sections, dry_run=args.dry_run)
            if not args.dry_run:
                update_campaign_html(cid, updated)
                print(f"\nCampaign ready: https://{SERVER}.admin.mailchimp.com/campaigns/edit?id={cid}")
        else:
            html = get_template_html(args.template_id)
            if not html:
                sys.exit("Template returned empty HTML. Try --duplicate to work via a campaign instead.")
            updated = inject_sections_into_html(html, sections, dry_run=args.dry_run)
            if not args.dry_run:
                update_template_html(args.template_id, updated)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
