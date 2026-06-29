"""
Mailchimp Template Updater
--------------------------
Reads a PDF or Word (.docx) monthly update file and updates a Mailchimp
template (or campaign) with the extracted content, preserving all formatting.

Usage:
    python mailchimp_update.py --file "May_26_Update.pdf" --template-id 10105010

    # Or update a campaign instead of a template:
    python mailchimp_update.py --file "May_26_Update.pdf" --campaign-id <uuid>

    # Duplicate a template first, then update the new campaign:
    python mailchimp_update.py --file "May_26_Update.pdf" --template-id 10105010 --duplicate

Requirements:
    pip install requests beautifulsoup4 pdfplumber python-docx python-dotenv lxml

Environment variables (set in .env or export):
    MAILCHIMP_API_KEY   your Mailchimp API key  (e.g. abc123-us3)
    MAILCHIMP_SERVER    your data center prefix  (e.g. us3)  — auto-detected from key
"""

import argparse
import os
import re
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

API_KEY = os.getenv("MAILCHIMP_API_KEY", "")
SERVER  = os.getenv("MAILCHIMP_SERVER", "")

if not SERVER and "-" in API_KEY:
    SERVER = API_KEY.split("-")[-1]  # e.g. "us3"

BASE_URL = f"https://{SERVER}.api.mailchimp.com/3.0"


def api(method: str, path: str, **kwargs):
    url = BASE_URL + path
    resp = requests.request(method, url, auth=("anystring", API_KEY), **kwargs)
    if not resp.ok:
        print(f"[ERROR] {method} {path} → {resp.status_code}: {resp.text}")
        sys.exit(1)
    return resp.json() if resp.text else {}


# ---------------------------------------------------------------------------
# PDF / DOCX extraction
# ---------------------------------------------------------------------------

def extract_text_from_pdf(path: str) -> str:
    try:
        import pdfplumber
    except ImportError:
        sys.exit("pdfplumber not installed. Run: pip install pdfplumber")

    text = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text.append(page.extract_text() or "")
    return "\n".join(text)


def extract_text_from_docx(path: str) -> str:
    try:
        from docx import Document
    except ImportError:
        sys.exit("python-docx not installed. Run: pip install python-docx")

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
    "executive_summary": r"Executive Summary\s*(.*?)(?=Operations|$)",
    "operations":        r"Operations\s*(.*?)(?=Revenue Initiatives|$)",
    "revenue":           r"Revenue Initiatives[:]*\s*(.*?)(?=Financials|$)",
    "financials":        r"Financials\s*(.*?)$",
}


def parse_sections(text: str) -> dict[str, str]:
    sections = {}
    for key, pattern in SECTION_PATTERNS.items():
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if match:
            content = match.group(1).strip()
            # Collapse excessive blank lines
            content = re.sub(r"\n{3,}", "\n\n", content)
            sections[key] = content
    return sections


def sections_to_html(sections: dict[str, str]) -> dict[str, str]:
    """Convert plain-text sections to minimal HTML paragraphs."""
    html_sections = {}
    for key, text in sections.items():
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        html_parts = []
        for para in paragraphs:
            lines = para.split("\n")
            if len(lines) == 1:
                html_parts.append(f"<p>{lines[0]}</p>")
            else:
                # Treat first line as a sub-heading if it looks like one
                first = lines[0]
                rest  = " ".join(lines[1:])
                if len(first) < 80 and not first.endswith("."):
                    html_parts.append(f"<p><strong>{first}</strong></p>")
                    if rest:
                        html_parts.append(f"<p>{rest}</p>")
                else:
                    html_parts.append(f"<p>{para.replace(chr(10), ' ')}</p>")
        html_sections[key] = "\n".join(html_parts)
    return html_sections


# ---------------------------------------------------------------------------
# Mailchimp template operations
# ---------------------------------------------------------------------------

def get_template(template_id: int) -> dict:
    print(f"Fetching template {template_id}...")
    return api("GET", f"/templates/{template_id}")


def get_template_html(template_id: int) -> str:
    data = get_template(template_id)
    # The template source is nested under 'active_source_url' or inline
    # For code templates the HTML is in data["active_source_url"]
    # For drag-and-drop it's available via the default_content endpoint
    default = api("GET", f"/templates/{template_id}/default-content")
    return default.get("html", "")


def find_editable_sections(html: str) -> list[str]:
    """Return the mc:edit region names found in the template HTML."""
    soup = BeautifulSoup(html, "lxml")
    regions = [tag["mc:edit"] for tag in soup.find_all(attrs={"mc:edit": True})]
    return regions


def replace_section_content(html: str, mc_edit_name: str, new_html: str) -> str:
    """Replace the inner HTML of the mc:edit region with new_html."""
    soup = BeautifulSoup(html, "lxml")
    tag = soup.find(attrs={"mc:edit": mc_edit_name})
    if not tag:
        print(f"  [WARN] mc:edit='{mc_edit_name}' not found in template — skipping")
        return html
    tag.clear()
    tag.append(BeautifulSoup(new_html, "lxml"))
    return str(soup)


def update_template(template_id: int, new_html: str, name: str = None):
    payload = {"html": new_html}
    if name:
        payload["name"] = name
    print(f"Updating template {template_id}...")
    api("PATCH", f"/templates/{template_id}", json=payload)
    print(f"  Template updated: https://{SERVER}.admin.mailchimp.com/templates/edit?id={template_id}")


# ---------------------------------------------------------------------------
# Campaign-based workflow (duplicate template → update campaign content)
# ---------------------------------------------------------------------------

def duplicate_template_as_campaign(template_id: int, subject: str, list_id: str) -> str:
    """Create a new campaign from a template. Returns campaign_id (UUID)."""
    print(f"Creating new campaign from template {template_id}...")
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
    campaign_id = resp["id"]
    print(f"  Campaign created: {campaign_id}")
    return campaign_id


def update_campaign_content_sections(campaign_id: str, sections: dict[str, str]):
    """Update named mc:edit sections in a campaign via the sections API."""
    print(f"Updating campaign {campaign_id} content sections...")
    payload = {"template": {"sections": sections}}
    api("PUT", f"/campaigns/{campaign_id}/content", json=payload)
    print(f"  Campaign content updated.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Update a Mailchimp template from a PDF/Word file.")
    parser.add_argument("--file",        required=True, help="Path to PDF or .docx update file")
    parser.add_argument("--template-id", type=int,      help="Mailchimp template ID to update directly")
    parser.add_argument("--campaign-id",                help="Mailchimp campaign UUID to update content in")
    parser.add_argument("--duplicate",   action="store_true",
                        help="Create a new campaign from the template (requires --template-id, --list-id, --subject)")
    parser.add_argument("--list-id",     help="Mailchimp audience/list ID (required with --duplicate)")
    parser.add_argument("--subject",     help="Email subject line (required with --duplicate)")
    parser.add_argument("--show-sections", action="store_true",
                        help="Print detected mc:edit section names and exit")
    args = parser.parse_args()

    if not API_KEY:
        sys.exit("Set MAILCHIMP_API_KEY in your environment or .env file.")

    # --- Extract content from file ---
    print(f"Reading {args.file}...")
    raw_text = extract_text(args.file)
    sections  = parse_sections(raw_text)
    html_secs = sections_to_html(sections)

    print(f"\nDetected sections: {list(sections.keys())}")
    for k, v in sections.items():
        preview = v[:120].replace("\n", " ")
        print(f"  [{k}] {preview}...")

    # --- Show mc:edit names if requested ---
    if args.show_sections and args.template_id:
        html = get_template_html(args.template_id)
        names = find_editable_sections(html)
        print(f"\nmc:edit regions in template {args.template_id}:")
        for n in names:
            print(f"  {n}")
        return

    # --- Update existing campaign ---
    if args.campaign_id:
        update_campaign_content_sections(args.campaign_id, html_secs)
        return

    # --- Duplicate template → new campaign ---
    if args.duplicate:
        if not args.template_id or not args.list_id or not args.subject:
            sys.exit("--duplicate requires --template-id, --list-id, and --subject")
        cid = duplicate_template_as_campaign(args.template_id, args.subject, args.list_id)
        update_campaign_content_sections(cid, html_secs)
        return

    # --- Update template HTML directly ---
    if args.template_id:
        html = get_template_html(args.template_id)
        if not html:
            print("[WARN] Could not fetch template HTML via default-content endpoint.")
            print("       The template may be a drag-and-drop type — use --campaign-id after duplicating manually.")
            sys.exit(1)

        # Map our section keys to mc:edit names in the template
        # Run with --show-sections first to discover the actual names
        section_map = {
            "executive_summary": "executive_summary",
            "operations":        "operations",
            "revenue":           "revenue_initiatives",
            "financials":        "financials",
        }

        for key, mc_name in section_map.items():
            if key in html_secs:
                html = replace_section_content(html, mc_name, html_secs[key])

        update_template(args.template_id, html)
        return

    parser.print_help()


if __name__ == "__main__":
    main()
