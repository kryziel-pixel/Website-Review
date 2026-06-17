"""
Weekly Property Review Automation
Checks links, addresses, Facebook activity, pricing, and specials
across property websites and apartments.com listings.
"""

import os
import re
import sys
import json
import asyncio
import argparse
import datetime
import urllib.parse
from dataclasses import dataclass, field
from typing import Optional

import requests
from bs4 import BeautifulSoup
from colorama import Fore, Style, init as colorama_init
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()
colorama_init(autoreset=True)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}
TIMEOUT = 15


# ── Status constants ──────────────────────────────────────────────────────────

PASS = "✓"
FAIL = "✗"
MANUAL = "Manual"
NA = "N/A"


@dataclass
class PropertyResult:
    name: str
    website: str
    apartments_com: str
    facebook: str

    # Item 1 – link checks
    link_book_tour: str = NA
    link_virtual_tour: str = NA
    link_application: str = NA
    link_social_media: str = NA
    link_menus: str = NA

    # Item 2 – address
    address_consistent: str = NA
    address_website: str = ""
    address_ils: str = ""
    address_note: str = ""

    # Item 3 – Facebook activity
    facebook_active: str = NA
    facebook_note: str = ""

    # Item 4 – pricing consistency
    pricing_consistent: str = NA
    pricing_website: str = ""
    pricing_ils: str = ""

    # Item 5 – specials displayed
    specials_on_ils: str = NA
    specials_on_website: str = NA

    # Item 6 – specials consistent
    specials_consistent: str = NA
    specials_ils_text: str = ""
    specials_website_text: str = ""

    # Item 7 – Google reviews
    google_rating: str = NA
    google_review_count: str = NA
    google_recent_feedback: str = MANUAL
    google_reviews_responded: str = MANUAL


# ── Helpers ───────────────────────────────────────────────────────────────────

def _playwright_fetch(url: str) -> Optional[str]:
    """Render a page with Playwright (headless Chromium) and return HTML."""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(
                ignore_https_errors=True,
                user_agent=HEADERS["User-Agent"],
                extra_http_headers={"Accept-Language": "en-US,en;q=0.9"},
            )
            page = ctx.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=30000)
            # Wait a bit for JS frameworks to render
            page.wait_for_timeout(2000)
            html = page.content()
            browser.close()
            return html
    except Exception as exc:
        print(f"    {Fore.YELLOW}Playwright failed for {url}: {exc}{Style.RESET_ALL}")
        return None


def fetch(url: str) -> Optional[BeautifulSoup]:
    """Fetch a page, trying requests first then Playwright as fallback."""
    # Try simple requests first (faster)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        if resp.status_code < 400 and len(resp.text) > 500:
            return BeautifulSoup(resp.text, "lxml")
    except Exception:
        pass

    # Fall back to Playwright for JS-heavy or bot-protected sites
    print(f"    {Fore.YELLOW}Falling back to Playwright for {url}{Style.RESET_ALL}")
    html = _playwright_fetch(url)
    if html:
        return BeautifulSoup(html, "lxml")
    return None


def check_url_alive(url: str) -> bool:
    try:
        r = requests.head(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        if r.status_code == 405:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        return r.status_code < 400
    except Exception:
        return False


def normalize_price(s: str) -> int:
    """Extract numeric value from price string like '$1,500'."""
    digits = re.sub(r"[^\d]", "", s)
    return int(digits) if digits else 0


def find_price_range(soup: BeautifulSoup) -> tuple[Optional[int], Optional[int]]:
    """Return (min_price, max_price) from a soup object."""
    text = soup.get_text(" ", strip=True)
    prices = re.findall(r"\$[\d,]+", text)
    values = [normalize_price(p) for p in prices if 500 <= normalize_price(p) <= 10000]
    if not values:
        return None, None
    return min(values), max(values)


def normalize_address(addr: str) -> str:
    addr = addr.lower().strip()
    addr = re.sub(r"\s+", " ", addr)
    for abbr, full in [("st", "street"), ("ave", "avenue"), ("blvd", "boulevard"),
                       ("dr", "drive"), ("rd", "road"), ("ln", "lane"), ("ct", "court")]:
        addr = re.sub(rf"\b{abbr}\b\.?", full, addr)
    return addr


def addresses_match(a: str, b: str) -> bool:
    if not a or not b:
        return False
    return normalize_address(a) == normalize_address(b)


def color_status(status: str) -> str:
    if status == PASS:
        return f"{Fore.GREEN}{status}{Style.RESET_ALL}"
    if status == FAIL:
        return f"{Fore.RED}{status}{Style.RESET_ALL}"
    if status == MANUAL:
        return f"{Fore.YELLOW}{status}{Style.RESET_ALL}"
    return status


# ── Item 1: Link Checks ───────────────────────────────────────────────────────

LINK_PATTERNS = {
    "book_tour": [
        r"book[\s\-_]?a[\s\-_]?tour", r"schedule[\s\-_]?a[\s\-_]?tour",
        r"request[\s\-_]?a[\s\-_]?tour", r"tour",
    ],
    "virtual_tour": [
        r"virtual[\s\-_]?tour", r"3d[\s\-_]?tour", r"matterport", r"floorplan.*tour",
    ],
    "application": [
        r"apply[\s\-_]?now", r"apply[\s\-_]?online", r"application",
        r"resident[\s\-_]?portal", r"renter[\s\-_]?portal",
    ],
    "social_media": [
        r"facebook\.com", r"instagram\.com", r"twitter\.com", r"x\.com",
        r"linkedin\.com", r"youtube\.com",
    ],
    "menus": [
        r"amenities", r"floor[\s\-_]?plans?", r"gallery", r"contact", r"about",
        r"photos", r"neighborhood",
    ],
}


def check_links(soup: BeautifulSoup, base_url: str) -> dict[str, str]:
    results = {}
    all_links = soup.find_all("a", href=True)

    def find_and_check(patterns: list[str], label: str) -> str:
        for a in all_links:
            href = a.get("href", "")
            text = a.get_text(" ", strip=True).lower()
            combined = (href + " " + text).lower()
            for pat in patterns:
                if re.search(pat, combined, re.IGNORECASE):
                    full_url = urllib.parse.urljoin(base_url, href)
                    alive = check_url_alive(full_url)
                    return PASS if alive else FAIL
        return FAIL

    results["book_tour"] = find_and_check(LINK_PATTERNS["book_tour"], "book_tour")
    results["virtual_tour"] = find_and_check(LINK_PATTERNS["virtual_tour"], "virtual_tour")
    results["application"] = find_and_check(LINK_PATTERNS["application"], "application")
    results["social_media"] = find_and_check(LINK_PATTERNS["social_media"], "social_media")
    results["menus"] = find_and_check(LINK_PATTERNS["menus"], "menus")
    return results


# ── Item 2: Address ───────────────────────────────────────────────────────────

ADDRESS_RE = re.compile(
    r"\d{2,5}\s+[A-Za-z0-9\s\.\-]+(?:Street|St|Avenue|Ave|Boulevard|Blvd|Drive|Dr|"
    r"Road|Rd|Lane|Ln|Court|Ct|Way|Place|Pl|Circle|Cir|Loop|Trail|Trl)[,\s]+[A-Za-z\s]+,\s*[A-Z]{2}\s*\d{5}",
    re.IGNORECASE,
)


def extract_address(soup: BeautifulSoup) -> str:
    if not soup:
        return ""
    for tag in ["footer", "address", "header"]:
        section = soup.find(tag)
        if section:
            text = section.get_text(" ", strip=True)
            m = ADDRESS_RE.search(text)
            if m:
                return m.group(0).strip()
    # Fall back to full page
    m = ADDRESS_RE.search(soup.get_text(" ", strip=True))
    return m.group(0).strip() if m else ""


def check_address(prop: dict, website_soup: BeautifulSoup, ils_soup: BeautifulSoup) -> tuple[str, str, str, str]:
    addr_web = extract_address(website_soup)
    addr_ils = extract_address(ils_soup)
    if not addr_web and not addr_ils:
        return MANUAL, addr_web, addr_ils, "Could not extract addresses automatically"
    if not addr_web or not addr_ils:
        return MANUAL, addr_web, addr_ils, "Could only extract one address"
    consistent = PASS if addresses_match(addr_web, addr_ils) else FAIL
    note = "" if consistent == PASS else f"Website: {addr_web} | ILS: {addr_ils}"
    return consistent, addr_web, addr_ils, note


# ── Item 3: Facebook Activity ─────────────────────────────────────────────────

def check_facebook(fb_url: str) -> tuple[str, str]:
    """
    Facebook heavily blocks scraping. We attempt a basic fetch and look for
    date signals. Most of the time this returns Manual.
    """
    try:
        resp = requests.get(fb_url, headers=HEADERS, timeout=TIMEOUT)
        if resp.status_code == 200:
            text = resp.text.lower()
            # If we got redirected to login wall
            if "log in" in text and "timeline" not in text:
                return MANUAL, "Facebook login required – check manually"
            # Look for date patterns suggesting recent posts
            today = datetime.date.today()
            week_ago = today - datetime.timedelta(days=7)
            dates_found = re.findall(r"(\d{4}-\d{2}-\d{2})", resp.text)
            recent = any(
                datetime.date.fromisoformat(d) >= week_ago
                for d in dates_found
                if _safe_date(d) is not None
            )
            if recent:
                return PASS, "Recent posts detected"
            if dates_found:
                return FAIL, f"No posts within last 7 days (last seen: {max(dates_found)})"
            return MANUAL, "Could not parse post dates – check manually"
    except Exception as exc:
        pass
    return MANUAL, f"Facebook blocked scraping – check manually: {fb_url}"


def _safe_date(s: str) -> Optional[datetime.date]:
    try:
        return datetime.date.fromisoformat(s)
    except ValueError:
        return None


# ── Item 4 & 5 & 6: Pricing and Specials ─────────────────────────────────────

SPECIAL_KEYWORDS = [
    r"weeks?\s+free", r"months?\s+free", r"move[\s\-]?in\s+special",
    r"concession", r"gift\s+card", r"reduced\s+rent", r"limited\s+time",
    r"summer\s+savings", r"look\s+&\s+lease", r"waived\s+fee",
    r"\d+\s*%\s+off", r"free\s+rent",
]
SPECIAL_RE = re.compile("|".join(SPECIAL_KEYWORDS), re.IGNORECASE)


def find_special_text(soup: BeautifulSoup) -> str:
    if not soup:
        return ""
    text = soup.get_text(" ", strip=True)
    matches = SPECIAL_RE.findall(text)
    if not matches:
        return ""
    # Return a short excerpt around the first match
    m = SPECIAL_RE.search(text)
    start = max(0, m.start() - 60)
    end = min(len(text), m.end() + 120)
    return text[start:end].strip()


def specials_similar(a: str, b: str) -> bool:
    """Rough similarity: check if key offer words overlap."""
    def tokens(s: str) -> set:
        return set(re.findall(r"\w+", s.lower())) - {"and", "or", "the", "a", "an", "in", "of"}
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return False
    overlap = len(ta & tb) / min(len(ta), len(tb))
    return overlap >= 0.4


def check_pricing(website_soup: BeautifulSoup, ils_soup: BeautifulSoup) -> tuple[str, str, str]:
    w_min, w_max = find_price_range(website_soup) if website_soup else (None, None)
    i_min, i_max = find_price_range(ils_soup) if ils_soup else (None, None)

    if None in (w_min, w_max, i_min, i_max):
        return MANUAL, f"${w_min}–${w_max}" if w_min else "N/A", f"${i_min}–${i_max}" if i_min else "N/A"

    w_str = f"${w_min:,}–${w_max:,}"
    i_str = f"${i_min:,}–${i_max:,}"
    threshold = 100  # dollars
    consistent = PASS if abs(w_min - i_min) <= threshold and abs(w_max - i_max) <= threshold else FAIL
    return consistent, w_str, i_str


# ── Item 7: Google Reviews ────────────────────────────────────────────────────

def get_google_reviews(place_name: str) -> tuple[str, str]:
    api_key = os.getenv("GOOGLE_API_KEY", "")
    if not api_key:
        return MANUAL, MANUAL

    try:
        # Find place
        search_url = "https://maps.googleapis.com/maps/api/place/findplacefromtext/json"
        params = {
            "input": place_name,
            "inputtype": "textquery",
            "fields": "place_id",
            "key": api_key,
        }
        r = requests.get(search_url, params=params, timeout=TIMEOUT)
        data = r.json()
        candidates = data.get("candidates", [])
        if not candidates:
            return MANUAL, MANUAL
        place_id = candidates[0]["place_id"]

        # Get details
        details_url = "https://maps.googleapis.com/maps/api/place/details/json"
        params = {
            "place_id": place_id,
            "fields": "rating,user_ratings_total",
            "key": api_key,
        }
        r = requests.get(details_url, params=params, timeout=TIMEOUT)
        detail = r.json().get("result", {})
        rating = str(detail.get("rating", MANUAL))
        count = str(detail.get("user_ratings_total", MANUAL))
        return rating, count
    except Exception as exc:
        print(f"    {Fore.YELLOW}Google API error: {exc}{Style.RESET_ALL}")
        return MANUAL, MANUAL


# ── Main Per-Property Review ──────────────────────────────────────────────────

def review_property(prop: dict) -> PropertyResult:
    name = prop["name"]
    print(f"\n{Fore.CYAN}{'='*60}")
    print(f"  Reviewing: {name}")
    print(f"{'='*60}{Style.RESET_ALL}")

    result = PropertyResult(
        name=name,
        website=prop["website"],
        apartments_com=prop["apartments_com"],
        facebook=prop["facebook"],
    )

    # Fetch pages
    print(f"  Fetching website...")
    website_soup = fetch(prop["website"])

    print(f"  Fetching apartments.com...")
    ils_soup = fetch(prop["apartments_com"])

    # Item 1: Links
    print(f"  Checking links (1.1–1.5)...")
    if website_soup:
        link_results = check_links(website_soup, prop["website"])
        result.link_book_tour = link_results["book_tour"]
        result.link_virtual_tour = link_results["virtual_tour"]
        result.link_application = link_results["application"]
        result.link_social_media = link_results["social_media"]
        result.link_menus = link_results["menus"]
    else:
        for attr in ["link_book_tour", "link_virtual_tour", "link_application",
                     "link_social_media", "link_menus"]:
            setattr(result, attr, FAIL)

    # Item 2: Address
    print(f"  Checking address (item 2)...")
    consistent, addr_web, addr_ils, note = check_address(prop, website_soup, ils_soup)
    result.address_consistent = consistent
    result.address_website = addr_web
    result.address_ils = addr_ils
    result.address_note = note + " | Google & Facebook: check manually"

    # Item 3: Facebook
    print(f"  Checking Facebook activity (item 3)...")
    fb_status, fb_note = check_facebook(prop["facebook"])
    result.facebook_active = fb_status
    result.facebook_note = fb_note

    # Item 4: Pricing
    print(f"  Checking pricing consistency (item 4)...")
    pricing_status, pricing_web, pricing_ils = check_pricing(website_soup, ils_soup)
    result.pricing_consistent = pricing_status
    result.pricing_website = pricing_web
    result.pricing_ils = pricing_ils

    # Item 5: Specials displayed
    print(f"  Checking specials (items 5–6)...")
    special_ils = find_special_text(ils_soup)
    special_web = find_special_text(website_soup)
    result.specials_on_ils = PASS if special_ils else FAIL
    result.specials_on_website = PASS if special_web else FAIL
    result.specials_ils_text = special_ils[:200] if special_ils else ""
    result.specials_website_text = special_web[:200] if special_web else ""

    # Item 6: Specials consistent
    if special_ils and special_web:
        result.specials_consistent = PASS if specials_similar(special_ils, special_web) else FAIL
    elif not special_ils and not special_web:
        result.specials_consistent = NA
    else:
        result.specials_consistent = FAIL

    # Item 7: Google reviews
    print(f"  Fetching Google reviews (item 7)...")
    rating, count = get_google_reviews(prop["google_place_name"])
    result.google_rating = rating
    result.google_review_count = count

    return result


# ── Console Report ────────────────────────────────────────────────────────────

def print_result(r: PropertyResult):
    def row(label: str, value: str, detail: str = ""):
        detail_str = f"  ({detail})" if detail else ""
        print(f"    {label:<40} {color_status(value)}{detail_str}")

    print(f"\n  {Fore.WHITE}{Style.BRIGHT}{r.name}{Style.RESET_ALL}")
    print(f"  {'-'*55}")
    row("1.1 Book a Tour", r.link_book_tour)
    row("1.2 Virtual Tour", r.link_virtual_tour)
    row("1.3 Application", r.link_application)
    row("1.4 Social Media Links", r.link_social_media)
    row("1.5 Menus", r.link_menus)
    row("2.  Address Consistent", r.address_consistent, r.address_note)
    row("3.  Facebook Active", r.facebook_active, r.facebook_note)
    row("4.  ILS/Website Pricing Consistent", r.pricing_consistent,
        f"Web: {r.pricing_website} | ILS: {r.pricing_ils}" if r.pricing_website else "")
    row("5a. Specials on ILS", r.specials_on_ils, r.specials_ils_text[:80])
    row("5b. Specials on Website", r.specials_on_website, r.specials_website_text[:80])
    row("6.  Specials Consistent", r.specials_consistent)
    row("7.1 Google Rating", r.google_rating)
    row("    Google Review Count", r.google_review_count)
    row("7.2 Recent Feedback", r.google_recent_feedback)
    row("7.3 Reviews Responded To", r.google_reviews_responded)


# ── Excel Report ──────────────────────────────────────────────────────────────

GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
HEADER_FILL = PatternFill("solid", fgColor="2E4057")
SUBHEADER_FILL = PatternFill("solid", fgColor="4A6FA5")
WHITE_FONT = Font(color="FFFFFF", bold=True)
BOLD_FONT = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


def status_fill(value: str) -> Optional[PatternFill]:
    if value == PASS:
        return GREEN_FILL
    if value == FAIL:
        return RED_FILL
    if value == MANUAL:
        return YELLOW_FILL
    return None


def write_excel(results: list[PropertyResult], filename: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly Review"

    today = datetime.date.today().strftime("%B %d, %Y")
    ws.merge_cells("A1:B1")
    ws["A1"] = f"As of {today}"
    ws["A1"].font = BOLD_FONT

    # Column headers
    ws["A2"] = "Item"
    ws["B2"] = "Things to Check"
    for col_idx, r in enumerate(results, start=3):
        cell = ws.cell(row=2, column=col_idx, value=r.name)
        cell.fill = HEADER_FILL
        cell.font = WHITE_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER

    ws["A2"].fill = HEADER_FILL
    ws["A2"].font = WHITE_FONT
    ws["A2"].alignment = CENTER
    ws["B2"].fill = HEADER_FILL
    ws["B2"].font = WHITE_FONT

    rows = [
        # (item_label, description, result_attr, is_section_header)
        ("1",   "Are the links working?", None, True),
        ("1.1", "Book a Tour", "link_book_tour", False),
        ("1.2", "Virtual Tour", "link_virtual_tour", False),
        ("1.3", "Application", "link_application", False),
        ("1.4", "Links to Social Media", "link_social_media", False),
        ("1.5", "Menus", "link_menus", False),
        ("2",   "Address all correct?", "address_consistent", False),
        ("3",   "Social Media Active?", "facebook_active", False),
        ("4",   "ILS and Website pricing consistent?", "pricing_consistent", False),
        ("5",   "Specials displayed in ILS and Website?", None, True),
        ("5a",  "Specials on ILS (apartments.com)", "specials_on_ils", False),
        ("5b",  "Specials on Website", "specials_on_website", False),
        ("6",   "Specials consistent in ILS and Website?", "specials_consistent", False),
        ("7",   "Google Reviews", None, True),
        ("7.1", "What is the current Google rating?", "google_rating", False),
        ("7.2", "Is recent feedback generally positive or negative?", "google_recent_feedback", False),
        ("7.3", "How many reviews were received last week?", "google_review_count", False),
        ("7.4", "Are reviews being responded to?", "google_reviews_responded", False),
    ]

    for row_idx, (item, desc, attr, is_header) in enumerate(rows, start=3):
        a_cell = ws.cell(row=row_idx, column=1, value=item)
        b_cell = ws.cell(row=row_idx, column=2, value=desc)

        if is_header:
            a_cell.fill = SUBHEADER_FILL
            a_cell.font = WHITE_FONT
            b_cell.fill = SUBHEADER_FILL
            b_cell.font = WHITE_FONT
        else:
            b_cell.font = BOLD_FONT if item in {"1", "2", "3", "4", "6"} else Font()

        a_cell.border = THIN_BORDER
        b_cell.border = THIN_BORDER

        for col_idx, r in enumerate(results, start=3):
            value = NA
            if attr:
                value = getattr(r, attr, NA)
            cell = ws.cell(row=row_idx, column=col_idx, value=value if attr else "")
            cell.alignment = CENTER
            cell.border = THIN_BORDER
            if attr:
                fill = status_fill(value)
                if fill:
                    cell.fill = fill

    # Column widths
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 42
    for col_idx in range(3, 3 + len(results)):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

    # Notes sheet
    ws_notes = wb.create_sheet("Notes & Details")
    ws_notes["A1"] = "Property"
    ws_notes["B1"] = "Item"
    ws_notes["C1"] = "Detail"
    for cell in [ws_notes["A1"], ws_notes["B1"], ws_notes["C1"]]:
        cell.fill = HEADER_FILL
        cell.font = WHITE_FONT

    note_row = 2
    for r in results:
        notes = [
            (r.name, "2. Address",
             f"Website: {r.address_website} | ILS: {r.address_ils} | {r.address_note}"),
            (r.name, "3. Facebook", r.facebook_note),
            (r.name, "4. Pricing",
             f"Website: {r.pricing_website} | ILS: {r.pricing_ils}"),
            (r.name, "5a. ILS Special", r.specials_ils_text),
            (r.name, "5b. Website Special", r.specials_website_text),
        ]
        for prop_name, item, detail in notes:
            if detail.strip():
                ws_notes.cell(row=note_row, column=1, value=prop_name)
                ws_notes.cell(row=note_row, column=2, value=item)
                ws_notes.cell(row=note_row, column=3, value=detail)
                note_row += 1

    ws_notes.column_dimensions["A"].width = 22
    ws_notes.column_dimensions["B"].width = 20
    ws_notes.column_dimensions["C"].width = 80

    wb.save(filename)
    print(f"\n{Fore.GREEN}Excel report saved: {filename}{Style.RESET_ALL}")


# ── Entry Point ───────────────────────────────────────────────────────────────

def main():
    from properties import PROPERTIES

    parser = argparse.ArgumentParser(description="Weekly property website review")
    parser.add_argument(
        "--property", "-p",
        help="Run for a single property by name (partial match, case-insensitive)",
    )
    parser.add_argument(
        "--no-excel", action="store_true",
        help="Skip Excel report generation",
    )
    args = parser.parse_args()

    props = PROPERTIES
    if args.property:
        props = [p for p in PROPERTIES if args.property.lower() in p["name"].lower()]
        if not props:
            print(f"{Fore.RED}No property matching '{args.property}' found.{Style.RESET_ALL}")
            print("Available:", ", ".join(p["name"] for p in PROPERTIES))
            sys.exit(1)

    print(f"\n{Fore.WHITE}{Style.BRIGHT}Weekly Property Review — {datetime.date.today()}{Style.RESET_ALL}")
    print(f"Reviewing {len(props)} propert{'y' if len(props)==1 else 'ies'}...\n")

    results = []
    for prop in props:
        try:
            result = review_property(prop)
            results.append(result)
            print_result(result)
        except Exception as exc:
            print(f"{Fore.RED}ERROR reviewing {prop['name']}: {exc}{Style.RESET_ALL}")

    if not args.no_excel and results:
        filename = f"report_{datetime.date.today().isoformat()}.xlsx"
        write_excel(results, filename)

    print(f"\n{Fore.WHITE}{Style.BRIGHT}Review complete.{Style.RESET_ALL}\n")


if __name__ == "__main__":
    main()
