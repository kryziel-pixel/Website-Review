"""
Weekly Property Review Automation
Checks links, addresses, Facebook activity, pricing, and specials
across property websites and apartments.com listings.
"""

import os
import re
import sys
import argparse
import datetime
import urllib.parse
from typing import Optional
from dataclasses import dataclass

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

PASS = "✓"
FAIL = "✗"
MANUAL = "Manual"
NA = "N/A"


# ── Result dataclass ──────────────────────────────────────────────────────────

@dataclass
class PropertyResult:
    name: str
    website: str
    apartments_com: str
    facebook: str

    link_book_tour: str = NA
    link_virtual_tour: str = NA
    link_application: str = NA
    link_social_media: str = NA
    link_menus: str = NA

    address_consistent: str = NA
    address_website: str = ""
    address_ils: str = ""
    address_note: str = ""

    facebook_active: str = NA
    facebook_note: str = ""

    pricing_consistent: str = NA
    pricing_website: str = ""
    pricing_ils: str = ""

    specials_on_ils: str = NA
    specials_on_website: str = NA

    specials_consistent: str = NA
    specials_ils_text: str = ""
    specials_website_text: str = ""

    google_rating: str = NA
    google_review_count: str = NA
    google_recent_feedback: str = MANUAL
    google_reviews_responded: str = MANUAL


# ── Playwright fetcher (primary) ──────────────────────────────────────────────

def fetch_with_playwright(url: str, wait_ms: int = 3000) -> Optional[str]:
    """Render a page with headless Chromium and return the full HTML."""
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
            page.wait_for_timeout(wait_ms)   # let JS frameworks finish rendering
            html = page.content()
            browser.close()
            return html
    except Exception as exc:
        print(f"    {Fore.YELLOW}Playwright failed for {url}: {exc}{Style.RESET_ALL}")
        return None


def fetch(url: str) -> Optional[BeautifulSoup]:
    """Fetch page HTML → BeautifulSoup. Uses Playwright so JS content is included."""
    # Try plain requests first (fast path for non-JS pages)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        if resp.status_code < 400 and len(resp.text) > 1000:
            soup = BeautifulSoup(resp.text, "lxml")
            page_text = soup.get_text()
            # If the page looks like a real site (has links/nav), use it
            if len(soup.find_all("a")) > 5:
                return soup
    except Exception:
        pass

    # Fall back to Playwright for JS-rendered sites
    print(f"    {Fore.YELLOW}Using Playwright for {url}{Style.RESET_ALL}")
    html = fetch_with_playwright(url)
    return BeautifulSoup(html, "lxml") if html else None


def check_url_alive(url: str) -> bool:
    try:
        r = requests.head(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        if r.status_code == 405:
            r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        return r.status_code < 400
    except Exception:
        return False


# ── Item 1: Link / Button Checks ──────────────────────────────────────────────

# Patterns to detect each link type in href or visible text
LINK_PATTERNS = {
    "book_tour": [
        r"book[\s\-_]?a[\s\-_]?tour", r"schedule[\s\-_]?a[\s\-_]?tour",
        r"request[\s\-_]?a[\s\-_]?tour", r"in[\s\-_]?person[\s\-_]?tour",
        r"self[\s\-_]?guided[\s\-_]?tour", r"come[\s\-_]?and[\s\-_]?tour",
    ],
    "virtual_tour": [
        r"virtual[\s\-_]?tour", r"3d[\s\-_]?tour", r"matterport",
        r"floorplan.*tour", r"tour.*3d",
    ],
    "application": [
        r"apply[\s\-_]?now", r"apply[\s\-_]?online", r"\bapplication\b",
        r"resident[\s\-_]?portal", r"renter[\s\-_]?portal",
        r"lease[\s\-_]?now", r"lease[\s\-_]?online", r"start[\s\-_]?application",
    ],
    "social_media": [
        r"facebook\.com", r"fb\.com", r"instagram\.com", r"twitter\.com",
        r"x\.com/(?!share)", r"linkedin\.com", r"youtube\.com",
    ],
    "menus": [
        r"\bamenities\b", r"floor[\s\-_]?plans?", r"\bgallery\b",
        r"\bcontact\b", r"\babout\b", r"\bphotos\b", r"\bneighborhood\b",
        r"\blifestyle\b", r"\bmap\b",
    ],
}


def _element_text_and_href(el) -> str:
    """Return combined href + text content of a tag for pattern matching."""
    href = el.get("href", "") or ""
    text = el.get_text(" ", strip=True)
    onclick = el.get("onclick", "") or ""
    aria = el.get("aria-label", "") or ""
    return (href + " " + text + " " + onclick + " " + aria).lower()


def check_links(soup: BeautifulSoup, base_url: str, has_virtual_tour: bool = True) -> dict[str, str]:
    results = {}

    # Collect both <a> and <button> elements — many sites use buttons for tour modals
    clickables = soup.find_all(["a", "button"])

    def find_match(patterns: list[str]) -> bool:
        for el in clickables:
            combined = _element_text_and_href(el)
            for pat in patterns:
                if re.search(pat, combined, re.IGNORECASE):
                    # For <a> tags verify the href is alive; buttons are always ✓
                    if el.name == "button":
                        return True
                    href = el.get("href", "")
                    if not href or href.startswith("#") or href.startswith("javascript"):
                        return True   # JS-triggered modal / anchor — assume works
                    full_url = urllib.parse.urljoin(base_url, href)
                    return check_url_alive(full_url)
        return False

    results["book_tour"] = PASS if find_match(LINK_PATTERNS["book_tour"]) else FAIL

    if not has_virtual_tour:
        results["virtual_tour"] = "N/A"
    else:
        results["virtual_tour"] = PASS if find_match(LINK_PATTERNS["virtual_tour"]) else FAIL

    results["application"] = PASS if find_match(LINK_PATTERNS["application"]) else FAIL
    results["social_media"] = PASS if find_match(LINK_PATTERNS["social_media"]) else FAIL
    results["menus"] = PASS if find_match(LINK_PATTERNS["menus"]) else FAIL

    return results


# ── Item 2: Address ───────────────────────────────────────────────────────────

ADDRESS_RE = re.compile(
    r"\d{2,5}\s+[A-Za-z0-9\s\.\-]+(?:Street|St|Avenue|Ave|Boulevard|Blvd|Drive|Dr|"
    r"Road|Rd|Lane|Ln|Court|Ct|Way|Place|Pl|Circle|Cir|Loop|Trail|Trl)[,\s]+"
    r"[A-Za-z\s]+,\s*[A-Z]{2}\s*\d{5}",
    re.IGNORECASE,
)


def extract_address(soup: BeautifulSoup) -> str:
    if not soup:
        return ""
    for tag in ["footer", "address", "header"]:
        section = soup.find(tag)
        if section:
            m = ADDRESS_RE.search(section.get_text(" ", strip=True))
            if m:
                return m.group(0).strip()
    m = ADDRESS_RE.search(soup.get_text(" ", strip=True))
    return m.group(0).strip() if m else ""


def normalize_address(addr: str) -> str:
    addr = addr.lower().strip()
    addr = re.sub(r"\s+", " ", addr)
    for abbr, full in [("st", "street"), ("ave", "avenue"), ("blvd", "boulevard"),
                       ("dr", "drive"), ("rd", "road"), ("ln", "lane"), ("ct", "court")]:
        addr = re.sub(rf"\b{abbr}\b\.?", full, addr)
    return addr


def check_address(website_soup, ils_soup) -> tuple[str, str, str, str]:
    addr_web = extract_address(website_soup)
    addr_ils = extract_address(ils_soup)
    if not addr_web and not addr_ils:
        return MANUAL, addr_web, addr_ils, "Could not extract addresses automatically"
    if not addr_web or not addr_ils:
        return MANUAL, addr_web, addr_ils, "Could only extract one address"
    consistent = PASS if normalize_address(addr_web) == normalize_address(addr_ils) else FAIL
    note = "" if consistent == PASS else f"Website: {addr_web} | ILS: {addr_ils}"
    return consistent, addr_web, addr_ils, note


# ── Item 3: Facebook Activity ─────────────────────────────────────────────────

def check_facebook(fb_url: str) -> tuple[str, str]:
    try:
        resp = requests.get(fb_url, headers=HEADERS, timeout=TIMEOUT)
        if resp.status_code == 200:
            text = resp.text.lower()
            if "log in" in text and "timeline" not in text:
                return MANUAL, "Facebook login required – check manually"
            today = datetime.date.today()
            week_ago = today - datetime.timedelta(days=7)
            dates_found = re.findall(r"(\d{4}-\d{2}-\d{2})", resp.text)
            recent = any(
                _safe_date(d) is not None and _safe_date(d) >= week_ago
                for d in dates_found
            )
            if recent:
                return PASS, "Recent posts detected"
            if dates_found:
                return FAIL, f"No posts within last 7 days"
    except Exception:
        pass
    return MANUAL, f"Facebook blocked scraping – check manually: {fb_url}"


def _safe_date(s: str) -> Optional[datetime.date]:
    try:
        return datetime.date.fromisoformat(s)
    except ValueError:
        return None


# ── Item 4: Pricing ───────────────────────────────────────────────────────────

def normalize_price(s: str) -> int:
    digits = re.sub(r"[^\d]", "", s)
    return int(digits) if digits else 0


def find_price_range(soup: BeautifulSoup) -> tuple[Optional[int], Optional[int]]:
    text = soup.get_text(" ", strip=True)
    prices = re.findall(r"\$[\d,]+", text)
    values = [normalize_price(p) for p in prices if 500 <= normalize_price(p) <= 10000]
    if not values:
        return None, None
    return min(values), max(values)


def check_pricing(website_soup, ils_soup) -> tuple[str, str, str]:
    w_min, w_max = find_price_range(website_soup) if website_soup else (None, None)
    i_min, i_max = find_price_range(ils_soup) if ils_soup else (None, None)
    if None in (w_min, w_max, i_min, i_max):
        w_str = f"${w_min:,}–${w_max:,}" if w_min else "N/A"
        i_str = f"${i_min:,}–${i_max:,}" if i_min else "N/A"
        return MANUAL, w_str, i_str
    w_str = f"${w_min:,}–${w_max:,}"
    i_str = f"${i_min:,}–${i_max:,}"
    threshold = 100
    consistent = PASS if abs(w_min - i_min) <= threshold and abs(w_max - i_max) <= threshold else FAIL
    return consistent, w_str, i_str


# ── Items 5 & 6: Specials ─────────────────────────────────────────────────────

SPECIAL_KEYWORDS = [
    r"\d+\s*weeks?\s+free",
    r"\d+\s*months?\s+free",
    r"weeks?\s+free\s+(?:base\s+)?rent",
    r"months?\s+free\s+(?:base\s+)?rent",
    r"move[\s\-]?in\s+special",
    r"pre[\s\-]?leasing\s+special",
    r"summer\s+savings",
    r"look\s*&\s*lease",
    r"gift\s+card",
    r"concession",
    r"reduced\s+rent",
    r"limited\s+time",
    r"waived\s+fee",
    r"\d+\s*%\s+off",
    r"free\s+rent",
    r"special\s+offer",
    r"leasing\s+special",
    r"rent\s+special",
]
SPECIAL_RE = re.compile("|".join(SPECIAL_KEYWORDS), re.IGNORECASE)


def find_special_text(soup: BeautifulSoup) -> str:
    if not soup:
        return ""
    # Check banner/announcement bar elements first (most prominent)
    for selector in ["[class*='banner']", "[class*='promo']", "[class*='special']",
                     "[class*='offer']", "[class*='alert']", "[class*='announcement']",
                     "header", "nav", ".hero", "#hero"]:
        try:
            els = soup.select(selector)
            for el in els:
                text = el.get_text(" ", strip=True)
                if SPECIAL_RE.search(text):
                    m = SPECIAL_RE.search(text)
                    start = max(0, m.start() - 40)
                    end = min(len(text), m.end() + 120)
                    return text[start:end].strip()
        except Exception:
            pass

    # Fall back to full page text
    text = soup.get_text(" ", strip=True)
    m = SPECIAL_RE.search(text)
    if m:
        start = max(0, m.start() - 40)
        end = min(len(text), m.end() + 120)
        return text[start:end].strip()
    return ""


def specials_similar(a: str, b: str) -> bool:
    def tokens(s: str) -> set:
        return set(re.findall(r"\w+", s.lower())) - {"and", "or", "the", "a", "an", "in", "of", "to"}
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return False
    return len(ta & tb) / min(len(ta), len(tb)) >= 0.4


# ── Item 7: Google Reviews ────────────────────────────────────────────────────

def get_google_reviews(place_name: str) -> tuple[str, str]:
    api_key = os.getenv("GOOGLE_API_KEY", "")
    if not api_key:
        return MANUAL, MANUAL
    try:
        r = requests.get(
            "https://maps.googleapis.com/maps/api/place/findplacefromtext/json",
            params={"input": place_name, "inputtype": "textquery",
                    "fields": "place_id", "key": api_key},
            timeout=TIMEOUT,
        )
        candidates = r.json().get("candidates", [])
        if not candidates:
            return MANUAL, MANUAL
        place_id = candidates[0]["place_id"]
        r2 = requests.get(
            "https://maps.googleapis.com/maps/api/place/details/json",
            params={"place_id": place_id,
                    "fields": "rating,user_ratings_total", "key": api_key},
            timeout=TIMEOUT,
        )
        detail = r2.json().get("result", {})
        return str(detail.get("rating", MANUAL)), str(detail.get("user_ratings_total", MANUAL))
    except Exception as exc:
        print(f"    {Fore.YELLOW}Google API error: {exc}{Style.RESET_ALL}")
        return MANUAL, MANUAL


# ── Main per-property review ──────────────────────────────────────────────────

def review_property(prop: dict) -> PropertyResult:
    name = prop["name"]
    has_virtual_tour = prop.get("has_virtual_tour", True)

    print(f"\n{Fore.CYAN}{'='*60}")
    print(f"  Reviewing: {name}")
    print(f"{'='*60}{Style.RESET_ALL}")

    result = PropertyResult(
        name=name,
        website=prop["website"],
        apartments_com=prop["apartments_com"],
        facebook=prop["facebook"],
    )

    print(f"  Fetching website (Playwright)...")
    website_html = fetch_with_playwright(prop["website"], wait_ms=4000)
    website_soup = BeautifulSoup(website_html, "lxml") if website_html else None

    print(f"  Fetching apartments.com (Playwright)...")
    ils_html = fetch_with_playwright(prop["apartments_com"], wait_ms=4000)
    ils_soup = BeautifulSoup(ils_html, "lxml") if ils_html else None

    # Item 1: Links & buttons
    print(f"  Checking links/buttons (1.1–1.5)...")
    if website_soup:
        lr = check_links(website_soup, prop["website"], has_virtual_tour)
        result.link_book_tour = lr["book_tour"]
        result.link_virtual_tour = lr["virtual_tour"]
        result.link_application = lr["application"]
        result.link_social_media = lr["social_media"]
        result.link_menus = lr["menus"]
    else:
        for attr in ["link_book_tour", "link_virtual_tour", "link_application",
                     "link_social_media", "link_menus"]:
            setattr(result, attr, FAIL)

    # Item 2: Address
    print(f"  Checking address (item 2)...")
    consistent, addr_web, addr_ils, note = check_address(website_soup, ils_soup)
    result.address_consistent = consistent
    result.address_website = addr_web
    result.address_ils = addr_ils
    result.address_note = (note + " | Google & Facebook: check manually").strip(" |")

    # Item 3: Facebook
    print(f"  Checking Facebook (item 3)...")
    fb_status, fb_note = check_facebook(prop["facebook"])
    result.facebook_active = fb_status
    result.facebook_note = fb_note

    # Item 4: Pricing
    print(f"  Checking pricing (item 4)...")
    pricing_status, pricing_web, pricing_ils = check_pricing(website_soup, ils_soup)
    result.pricing_consistent = pricing_status
    result.pricing_website = pricing_web
    result.pricing_ils = pricing_ils

    # Items 5 & 6: Specials
    print(f"  Checking specials (items 5–6)...")
    special_ils = find_special_text(ils_soup) if ils_soup else ""
    special_web = find_special_text(website_soup) if website_soup else ""
    result.specials_on_ils = PASS if special_ils else FAIL
    result.specials_on_website = PASS if special_web else FAIL
    result.specials_ils_text = special_ils[:250]
    result.specials_website_text = special_web[:250]

    if special_ils and special_web:
        result.specials_consistent = PASS if specials_similar(special_ils, special_web) else FAIL
    elif not special_ils and not special_web:
        result.specials_consistent = NA
    else:
        result.specials_consistent = FAIL

    # Item 7: Google
    print(f"  Fetching Google reviews (item 7)...")
    result.google_rating, result.google_review_count = get_google_reviews(prop["google_place_name"])

    return result


# ── Console output ────────────────────────────────────────────────────────────

def color_status(value: str) -> str:
    if value == PASS:
        return f"{Fore.GREEN}{value}{Style.RESET_ALL}"
    if value == FAIL:
        return f"{Fore.RED}{value}{Style.RESET_ALL}"
    if value == MANUAL:
        return f"{Fore.YELLOW}{value}{Style.RESET_ALL}"
    return value


def print_result(r: PropertyResult):
    def row(label, value, detail=""):
        d = f"  ({detail})" if detail else ""
        print(f"    {label:<42} {color_status(value)}{d}")

    print(f"\n  {Fore.WHITE}{Style.BRIGHT}{r.name}{Style.RESET_ALL}")
    print(f"  {'-'*56}")
    row("1.1 Book a Tour", r.link_book_tour)
    row("1.2 Virtual Tour", r.link_virtual_tour)
    row("1.3 Application", r.link_application)
    row("1.4 Social Media Links", r.link_social_media)
    row("1.5 Menus", r.link_menus)
    row("2.  Address Consistent", r.address_consistent, r.address_note)
    row("3.  Facebook Active", r.facebook_active, r.facebook_note)
    row("4.  ILS/Website Pricing",  r.pricing_consistent,
        f"Web: {r.pricing_website} | ILS: {r.pricing_ils}" if r.pricing_website else "")
    row("5a. Specials on ILS", r.specials_on_ils, r.specials_ils_text[:80])
    row("5b. Specials on Website", r.specials_on_website, r.specials_website_text[:80])
    row("6.  Specials Consistent", r.specials_consistent)
    row("7.1 Google Rating", r.google_rating)
    row("    Google Review Count", r.google_review_count)
    row("7.2 Recent Feedback", r.google_recent_feedback)
    row("7.4 Reviews Responded To", r.google_reviews_responded)


# ── Excel report ──────────────────────────────────────────────────────────────

GREEN_FILL   = PatternFill("solid", fgColor="C6EFCE")
RED_FILL     = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL  = PatternFill("solid", fgColor="FFEB9C")
HEADER_FILL  = PatternFill("solid", fgColor="2E4057")
SUBHDR_FILL  = PatternFill("solid", fgColor="4A6FA5")
WHITE_FONT   = Font(color="FFFFFF", bold=True)
BOLD_FONT    = Font(bold=True)
CENTER       = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER  = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

ROWS = [
    ("1",   "Are the links working?",                         None,                    True),
    ("1.1", "Book a Tour",                                    "link_book_tour",         False),
    ("1.2", "Virtual Tour",                                   "link_virtual_tour",      False),
    ("1.3", "Application",                                    "link_application",       False),
    ("1.4", "Links to Social Media",                          "link_social_media",      False),
    ("1.5", "Menus",                                          "link_menus",             False),
    ("2",   "Address all correct?",                           "address_consistent",     False),
    ("3",   "Social Media Active?",                           "facebook_active",        False),
    ("4",   "ILS and Website pricing consistent?",            "pricing_consistent",     False),
    ("5",   "Specials displayed in ILS and Website?",         None,                    True),
    ("5a",  "Specials on ILS (apartments.com)",               "specials_on_ils",        False),
    ("5b",  "Specials on Website",                            "specials_on_website",    False),
    ("6",   "Specials consistent in ILS and Website?",        "specials_consistent",   False),
    ("7",   "Google Reviews",                                 None,                    True),
    ("7.1", "What is the current Google rating?",             "google_rating",          False),
    ("7.2", "Is recent feedback generally positive/negative?","google_recent_feedback", False),
    ("7.3", "How many reviews received last week?",           "google_review_count",    False),
    ("7.4", "Are reviews being responded to?",                "google_reviews_responded",False),
]


def _status_fill(value: str) -> Optional[PatternFill]:
    if value == PASS:   return GREEN_FILL
    if value == FAIL:   return RED_FILL
    if value == MANUAL: return YELLOW_FILL
    return None


def write_excel(results: list[PropertyResult], filename: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Weekly Review"

    today_str = datetime.date.today().strftime("%B %d, %Y")
    ws.merge_cells("A1:B1")
    ws["A1"] = f"As of {today_str}"
    ws["A1"].font = BOLD_FONT

    # Header row
    ws["A2"] = "Item"
    ws["B2"] = "Things to Check"
    for c, r in enumerate(results, start=3):
        cell = ws.cell(row=2, column=c, value=r.name)
        cell.fill = HEADER_FILL
        cell.font = WHITE_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    for col in ["A2", "B2"]:
        ws[col].fill = HEADER_FILL
        ws[col].font = WHITE_FONT
        ws[col].alignment = CENTER
        ws[col].border = THIN_BORDER

    # Data rows
    for ri, (item, desc, attr, is_header) in enumerate(ROWS, start=3):
        a = ws.cell(row=ri, column=1, value=item)
        b = ws.cell(row=ri, column=2, value=desc)
        if is_header:
            a.fill = SUBHDR_FILL; a.font = WHITE_FONT
            b.fill = SUBHDR_FILL; b.font = WHITE_FONT
        else:
            b.font = BOLD_FONT
        a.border = THIN_BORDER
        b.border = THIN_BORDER

        for ci, r in enumerate(results, start=3):
            val = getattr(r, attr, "") if attr else ""
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = CENTER
            cell.border = THIN_BORDER
            if attr:
                fill = _status_fill(val)
                if fill:
                    cell.fill = fill

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 44
    for ci in range(3, 3 + len(results)):
        ws.column_dimensions[get_column_letter(ci)].width = 18

    # Notes sheet
    ws2 = wb.create_sheet("Notes & Details")
    for ci, hdr in enumerate(["Property", "Item", "Detail"], start=1):
        c = ws2.cell(row=1, column=ci, value=hdr)
        c.fill = HEADER_FILL
        c.font = WHITE_FONT
    nr = 2
    for r in results:
        notes = [
            (r.name, "2. Address",          f"Website: {r.address_website} | ILS: {r.address_ils} | {r.address_note}"),
            (r.name, "3. Facebook",         r.facebook_note),
            (r.name, "4. Pricing",          f"Website: {r.pricing_website} | ILS: {r.pricing_ils}"),
            (r.name, "5a. ILS Special",     r.specials_ils_text),
            (r.name, "5b. Website Special", r.specials_website_text),
        ]
        for pname, itm, detail in notes:
            if detail.strip():
                ws2.cell(row=nr, column=1, value=pname)
                ws2.cell(row=nr, column=2, value=itm)
                ws2.cell(row=nr, column=3, value=detail)
                nr += 1
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 20
    ws2.column_dimensions["C"].width = 90

    wb.save(filename)
    print(f"\n{Fore.GREEN}Excel report saved: {filename}{Style.RESET_ALL}")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    from properties import PROPERTIES

    parser = argparse.ArgumentParser(description="Weekly property website review")
    parser.add_argument("--property", "-p",
                        help="Single property name (partial match, case-insensitive)")
    parser.add_argument("--no-excel", action="store_true",
                        help="Skip Excel report")
    args = parser.parse_args()

    props = PROPERTIES
    if args.property:
        props = [p for p in PROPERTIES if args.property.lower() in p["name"].lower()]
        if not props:
            print(f"{Fore.RED}No property matching '{args.property}'.{Style.RESET_ALL}")
            print("Available:", ", ".join(p["name"] for p in PROPERTIES))
            sys.exit(1)

    print(f"\n{Fore.WHITE}{Style.BRIGHT}Weekly Property Review — {datetime.date.today()}{Style.RESET_ALL}")
    print(f"Reviewing {len(props)} propert{'y' if len(props)==1 else 'ies'}...\n")
    print(f"{Fore.YELLOW}Note: Uses Playwright (headless browser) — expect ~2 min per property.{Style.RESET_ALL}\n")

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
