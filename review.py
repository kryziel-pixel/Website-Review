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

TIMEOUT = 15

PASS   = "✓"
FAIL   = "✗"
MANUAL = "Manual"
NA     = "N/A"


# ── Result dataclass ──────────────────────────────────────────────────────────

@dataclass
class PropertyResult:
    name: str
    website: str
    apartments_com: str
    facebook: str

    link_book_tour: str    = NA
    link_virtual_tour: str = NA
    link_application: str  = NA
    link_social_media: str = NA
    link_menus: str        = NA

    address_consistent: str = NA
    address_website: str    = ""
    address_ils: str        = ""
    address_note: str       = ""

    facebook_active: str = NA
    facebook_note: str   = ""

    pricing_consistent: str = NA
    pricing_website: str    = ""
    pricing_ils: str        = ""

    specials_on_ils: str     = NA
    specials_on_website: str = NA

    specials_consistent: str   = NA
    specials_ils_text: str     = ""
    specials_website_text: str = ""

    google_rating: str        = NA
    google_review_count: str  = NA
    google_recent_feedback: str  = MANUAL
    google_reviews_responded: str = MANUAL


# ── Playwright fetcher ────────────────────────────────────────────────────────

def fetch_with_playwright(url: str, wait_ms: int = 4000, scroll: bool = True) -> Optional[str]:
    """
    Render a page with headless Chromium.
    - Spoofs common bot-detection signals
    - Scrolls to the bottom so lazy-loaded footer/widgets appear
    """
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                ],
            )
            ctx = browser.new_context(
                ignore_https_errors=True,
                user_agent=(
                    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1280, "height": 900},
                extra_http_headers={
                    "Accept-Language": "en-US,en;q=0.9",
                    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                },
            )
            # Hide webdriver flag that sites use to detect bots
            ctx.add_init_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
            page = ctx.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=40000)
            page.wait_for_timeout(wait_ms)

            if scroll:
                # Scroll gradually to trigger lazy-loaded content (footer, widgets)
                page.evaluate("""
                    async () => {
                        await new Promise(resolve => {
                            let y = 0;
                            const step = 300;
                            const delay = 150;
                            const timer = setInterval(() => {
                                window.scrollBy(0, step);
                                y += step;
                                if (y >= document.body.scrollHeight) {
                                    clearInterval(timer);
                                    window.scrollTo(0, 0);
                                    resolve();
                                }
                            }, delay);
                        });
                    }
                """)
                page.wait_for_timeout(2000)  # wait for any lazy content to render

            html = page.content()
            browser.close()
            return html
    except Exception as exc:
        print(f"    {Fore.YELLOW}Playwright failed for {url}: {exc}{Style.RESET_ALL}")
        return None


def fetch(url: str) -> Optional[BeautifulSoup]:
    html = fetch_with_playwright(url)
    return BeautifulSoup(html, "lxml") if html else None


# ── Item 1: Link / Button / Widget checks ────────────────────────────────────

# Third-party platforms whose presence on the page signals a feature works
BOOKING_PLATFORMS  = ["quext", "rentcafe", "funnel", "knock", "appfolio",
                      "yardi", "realpage", "entrata", "resman", "leasestar",
                      "scheduleatour", "tourschedule", "selfguidedtour"]
APPLY_PLATFORMS    = ["rentcafe", "appfolio", "yardi", "realpage", "entrata",
                      "resman", "leasestar", "mynd", "buildium", "cozy"]
SOCIAL_DOMAINS     = ["facebook.com", "fb.com", "instagram.com", "twitter.com",
                      "x.com", "linkedin.com", "youtube.com", "tiktok.com"]
VIRTUAL_TOUR_SRCS  = ["matterport", "kuula", "3dvista", "roundme", "vtour",
                      "virtualtour", "virtual-tour", "3d-tour", "iguide"]

BOOK_TOUR_TEXT_RE = re.compile(
    r"book\s*a?\s*tour|schedule\s*a?\s*tour|request\s*a?\s*tour|"
    r"in[\s\-]?person\s*tour|self[\s\-]?guided|come\s*and\s*tour|"
    r"tour\s*now|tour\s*today",
    re.IGNORECASE,
)
APPLY_TEXT_RE = re.compile(
    r"apply\s*now|apply\s*online|lease\s*now|lease\s*online|"
    r"start\s*application|apply\s*today|application\b|resident\s*portal",
    re.IGNORECASE,
)
VIRTUAL_TOUR_TEXT_RE = re.compile(
    r"virtual\s*tour|3d\s*tour|take\s*a\s*tour\s*online|tour\s*from\s*home",
    re.IGNORECASE,
)
MENU_TEXT_RE = re.compile(
    r"\bamenities\b|\bfloor\s*plans?\b|\bgallery\b|\bcontact\b|\bphotos?\b|"
    r"\bneighborhood\b|\blifestyle\b|\bmap\b|\babout\b",
    re.IGNORECASE,
)


def _page_contains(soup: BeautifulSoup, text_re, platform_list: list[str] = None) -> bool:
    """
    Return True if any <a>, <button>, script src, or iframe src matches.
    Checks visible text, href, onclick, aria-label, and external script/iframe sources.
    """
    full_html = str(soup).lower()

    # Check script/iframe sources for third-party platforms
    if platform_list:
        for src_tag in soup.find_all(["script", "iframe"], src=True):
            src = (src_tag.get("src") or "").lower()
            if any(p in src for p in platform_list):
                return True
        # Also check raw HTML for platform names (some load via JS variables)
        if any(p in full_html for p in platform_list):
            return True

    # Check all clickable elements
    for el in soup.find_all(["a", "button", "div", "span"]):
        text  = el.get_text(" ", strip=True)
        href  = el.get("href", "") or ""
        aria  = el.get("aria-label", "") or ""
        onclick = el.get("onclick", "") or ""
        combined = " ".join([text, href, aria, onclick])
        if text_re and text_re.search(combined):
            return True

    return False


def check_links(soup: BeautifulSoup, base_url: str, has_virtual_tour: bool = True) -> dict[str, str]:
    results = {}

    results["book_tour"] = (
        PASS if _page_contains(soup, BOOK_TOUR_TEXT_RE, BOOKING_PLATFORMS) else FAIL
    )

    if not has_virtual_tour:
        results["virtual_tour"] = "N/A"
    else:
        results["virtual_tour"] = (
            PASS if _page_contains(soup, VIRTUAL_TOUR_TEXT_RE, VIRTUAL_TOUR_SRCS) else FAIL
        )

    results["application"] = (
        PASS if _page_contains(soup, APPLY_TEXT_RE, APPLY_PLATFORMS) else FAIL
    )

    # Social media: check raw HTML for social domain names (fastest, most reliable)
    full_html = str(soup).lower()
    results["social_media"] = (
        PASS if any(domain in full_html for domain in SOCIAL_DOMAINS) else FAIL
    )

    results["menus"] = (
        PASS if _page_contains(soup, MENU_TEXT_RE) else FAIL
    )

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
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "en-US,en;q=0.9",
        }
        resp = requests.get(fb_url, headers=headers, timeout=TIMEOUT)
        if resp.status_code == 200:
            text = resp.text.lower()
            if "log in" in text and "timeline" not in text:
                return MANUAL, "Facebook login required – check manually"
            today     = datetime.date.today()
            week_ago  = today - datetime.timedelta(days=7)
            dates     = re.findall(r"(\d{4}-\d{2}-\d{2})", resp.text)
            if any(_safe_date(d) and _safe_date(d) >= week_ago for d in dates):
                return PASS, "Recent posts detected"
            if dates:
                return FAIL, "No posts within last 7 days"
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
    text   = soup.get_text(" ", strip=True)
    prices = re.findall(r"\$[\d,]+", text)
    values = [normalize_price(p) for p in prices if 500 <= normalize_price(p) <= 10000]
    if not values:
        return None, None
    return min(values), max(values)


def check_pricing(website_soup, ils_soup) -> tuple[str, str, str]:
    w_min, w_max = find_price_range(website_soup) if website_soup else (None, None)
    i_min, i_max = find_price_range(ils_soup)     if ils_soup     else (None, None)
    w_str = f"${w_min:,}–${w_max:,}" if w_min else "N/A"
    i_str = f"${i_min:,}–${i_max:,}" if i_min else "N/A"
    if None in (w_min, w_max, i_min, i_max):
        return MANUAL, w_str, i_str
    consistent = PASS if abs(w_min - i_min) <= 100 and abs(w_max - i_max) <= 100 else FAIL
    return consistent, w_str, i_str


# ── Items 5 & 6: Specials ─────────────────────────────────────────────────────

SPECIAL_RE = re.compile(
    r"\d+\s*weeks?\s+free|"
    r"\d+\s*months?\s+free|"
    r"weeks?\s+free\s+(?:base\s+)?rent|"
    r"months?\s+free\s+(?:base\s+)?rent|"
    r"move[\s\-]?in\s+special|"
    r"pre[\s\-]?leasing\s+special|"
    r"leasing\s+special|"
    r"rent\s+special|"
    r"summer\s+savings|"
    r"look\s*&\s*lease|"
    r"gift\s+card|"
    r"concession|"
    r"reduced\s+rent|"
    r"waived\s+(?:admin\s+)?fee|"
    r"\d+\s*%\s+off\s+rent|"
    r"free\s+(?:base\s+)?rent|"
    r"special\s+offer",
    re.IGNORECASE,
)


def find_special_text(soup: BeautifulSoup) -> str:
    if not soup:
        return ""
    # Priority: check banner/announcement/promo elements first
    priority_selectors = [
        "[class*='banner']", "[class*='promo']", "[class*='special']",
        "[class*='offer']",  "[class*='alert']", "[class*='announcement']",
        "[class*='notice']", "[class*='ribbon']", "[class*='hero']",
        "header", "nav",
    ]
    for sel in priority_selectors:
        try:
            for el in soup.select(sel):
                text = el.get_text(" ", strip=True)
                m = SPECIAL_RE.search(text)
                if m:
                    s = max(0, m.start() - 40)
                    e = min(len(text), m.end() + 150)
                    return text[s:e].strip()
        except Exception:
            pass

    # Fall back to full page
    text = soup.get_text(" ", strip=True)
    m = SPECIAL_RE.search(text)
    if m:
        s = max(0, m.start() - 40)
        e = min(len(text), m.end() + 150)
        return text[s:e].strip()
    return ""


def specials_similar(a: str, b: str) -> bool:
    stop = {"and", "or", "the", "a", "an", "in", "of", "to", "on", "at", "for"}
    def tokens(s):
        return set(re.findall(r"\w+", s.lower())) - stop
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
    name             = prop["name"]
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

    print(f"  Fetching website (Playwright + scroll)...")
    website_html = fetch_with_playwright(prop["website"], wait_ms=5000, scroll=True)
    website_soup = BeautifulSoup(website_html, "lxml") if website_html else None

    print(f"  Fetching apartments.com (Playwright + stealth)...")
    ils_html = fetch_with_playwright(prop["apartments_com"], wait_ms=5000, scroll=True)
    ils_soup = BeautifulSoup(ils_html, "lxml") if ils_html else None

    # Item 1
    print(f"  Checking links/buttons (1.1–1.5)...")
    if website_soup:
        lr = check_links(website_soup, prop["website"], has_virtual_tour)
        result.link_book_tour    = lr["book_tour"]
        result.link_virtual_tour = lr["virtual_tour"]
        result.link_application  = lr["application"]
        result.link_social_media = lr["social_media"]
        result.link_menus        = lr["menus"]
    else:
        for attr in ["link_book_tour","link_virtual_tour","link_application",
                     "link_social_media","link_menus"]:
            setattr(result, attr, FAIL)

    # Item 2
    print(f"  Checking address (item 2)...")
    consistent, addr_web, addr_ils, note = check_address(website_soup, ils_soup)
    result.address_consistent = consistent
    result.address_website    = addr_web
    result.address_ils        = addr_ils
    result.address_note       = (note + " | Google & Facebook: check manually").strip(" |")

    # Item 3
    print(f"  Checking Facebook (item 3)...")
    result.facebook_active, result.facebook_note = check_facebook(prop["facebook"])

    # Item 4
    print(f"  Checking pricing (item 4)...")
    result.pricing_consistent, result.pricing_website, result.pricing_ils = (
        check_pricing(website_soup, ils_soup)
    )

    # Items 5 & 6
    print(f"  Checking specials (items 5–6)...")
    special_ils = find_special_text(ils_soup)   if ils_soup     else ""
    special_web = find_special_text(website_soup) if website_soup else ""
    result.specials_on_ils      = PASS if special_ils else FAIL
    result.specials_on_website  = PASS if special_web else FAIL
    result.specials_ils_text    = special_ils[:250]
    result.specials_website_text = special_web[:250]
    if special_ils and special_web:
        result.specials_consistent = PASS if specials_similar(special_ils, special_web) else FAIL
    elif not special_ils and not special_web:
        result.specials_consistent = NA
    else:
        result.specials_consistent = FAIL

    # Item 7
    print(f"  Fetching Google reviews (item 7)...")
    result.google_rating, result.google_review_count = (
        get_google_reviews(prop["google_place_name"])
    )

    return result


# ── Console output ────────────────────────────────────────────────────────────

def color_status(value: str) -> str:
    if value == PASS:   return f"{Fore.GREEN}{value}{Style.RESET_ALL}"
    if value == FAIL:   return f"{Fore.RED}{value}{Style.RESET_ALL}"
    if value == MANUAL: return f"{Fore.YELLOW}{value}{Style.RESET_ALL}"
    return value


def print_result(r: PropertyResult):
    def row(label, value, detail=""):
        d = f"  ({detail})" if detail else ""
        print(f"    {label:<42} {color_status(value)}{d}")

    print(f"\n  {Fore.WHITE}{Style.BRIGHT}{r.name}{Style.RESET_ALL}")
    print(f"  {'-'*56}")
    row("1.1 Book a Tour",          r.link_book_tour)
    row("1.2 Virtual Tour",         r.link_virtual_tour)
    row("1.3 Application",          r.link_application)
    row("1.4 Social Media Links",   r.link_social_media)
    row("1.5 Menus",                r.link_menus)
    row("2.  Address Consistent",   r.address_consistent, r.address_note)
    row("3.  Facebook Active",      r.facebook_active,    r.facebook_note)
    row("4.  ILS/Website Pricing",  r.pricing_consistent,
        f"Web: {r.pricing_website} | ILS: {r.pricing_ils}" if r.pricing_website else "")
    row("5a. Specials on ILS",      r.specials_on_ils,      r.specials_ils_text[:80])
    row("5b. Specials on Website",  r.specials_on_website,  r.specials_website_text[:80])
    row("6.  Specials Consistent",  r.specials_consistent)
    row("7.1 Google Rating",        r.google_rating)
    row("    Google Review Count",  r.google_review_count)
    row("7.2 Recent Feedback",      r.google_recent_feedback)
    row("7.4 Reviews Responded To", r.google_reviews_responded)


# ── Excel report ──────────────────────────────────────────────────────────────

GREEN_FILL  = PatternFill("solid", fgColor="C6EFCE")
RED_FILL    = PatternFill("solid", fgColor="FFC7CE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
HEADER_FILL = PatternFill("solid", fgColor="2E4057")
SUBHDR_FILL = PatternFill("solid", fgColor="4A6FA5")
WHITE_FONT  = Font(color="FFFFFF", bold=True)
BOLD_FONT   = Font(bold=True)
CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

ROWS = [
    ("1",   "Are the links working?",                          None,                     True),
    ("1.1", "Book a Tour",                                     "link_book_tour",          False),
    ("1.2", "Virtual Tour",                                    "link_virtual_tour",       False),
    ("1.3", "Application",                                     "link_application",        False),
    ("1.4", "Links to Social Media",                           "link_social_media",       False),
    ("1.5", "Menus",                                           "link_menus",              False),
    ("2",   "Address all correct?",                            "address_consistent",      False),
    ("3",   "Social Media Active?",                            "facebook_active",         False),
    ("4",   "ILS and Website pricing consistent?",             "pricing_consistent",      False),
    ("5",   "Specials displayed in ILS and Website?",          None,                      True),
    ("5a",  "Specials on ILS (apartments.com)",                "specials_on_ils",         False),
    ("5b",  "Specials on Website",                             "specials_on_website",     False),
    ("6",   "Specials consistent in ILS and Website?",         "specials_consistent",     False),
    ("7",   "Google Reviews",                                  None,                      True),
    ("7.1", "What is the current Google rating?",              "google_rating",           False),
    ("7.2", "Is recent feedback generally positive/negative?", "google_recent_feedback",  False),
    ("7.3", "How many reviews received last week?",            "google_review_count",     False),
    ("7.4", "Are reviews being responded to?",                 "google_reviews_responded",False),
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

    ws.merge_cells("A1:B1")
    ws["A1"] = f"As of {datetime.date.today().strftime('%B %d, %Y')}"
    ws["A1"].font = BOLD_FONT

    for col, hdr in [("A2","Item"), ("B2","Things to Check")]:
        ws[col] = hdr
        ws[col].fill = HEADER_FILL
        ws[col].font = WHITE_FONT
        ws[col].alignment = CENTER
        ws[col].border = THIN_BORDER

    for c, r in enumerate(results, start=3):
        cell = ws.cell(row=2, column=c, value=r.name)
        cell.fill = HEADER_FILL; cell.font = WHITE_FONT
        cell.alignment = CENTER;  cell.border = THIN_BORDER

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
            val  = getattr(r, attr, "") if attr else ""
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = CENTER; cell.border = THIN_BORDER
            if attr:
                fill = _status_fill(val)
                if fill:
                    cell.fill = fill

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 44
    for ci in range(3, 3 + len(results)):
        ws.column_dimensions[get_column_letter(ci)].width = 18

    ws2 = wb.create_sheet("Notes & Details")
    for ci, hdr in enumerate(["Property","Item","Detail"], start=1):
        c = ws2.cell(row=1, column=ci, value=hdr)
        c.fill = HEADER_FILL; c.font = WHITE_FONT
    nr = 2
    for r in results:
        for pname, itm, detail in [
            (r.name, "2. Address",          f"Website: {r.address_website} | ILS: {r.address_ils} | {r.address_note}"),
            (r.name, "3. Facebook",         r.facebook_note),
            (r.name, "4. Pricing",          f"Website: {r.pricing_website} | ILS: {r.pricing_ils}"),
            (r.name, "5a. ILS Special",     r.specials_ils_text),
            (r.name, "5b. Website Special", r.specials_website_text),
        ]:
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
    parser.add_argument("--no-excel", action="store_true")
    args = parser.parse_args()

    props = PROPERTIES
    if args.property:
        props = [p for p in PROPERTIES if args.property.lower() in p["name"].lower()]
        if not props:
            print(f"{Fore.RED}No property matching '{args.property}'.{Style.RESET_ALL}")
            print("Available:", ", ".join(p["name"] for p in PROPERTIES))
            sys.exit(1)

    print(f"\n{Fore.WHITE}{Style.BRIGHT}Weekly Property Review — {datetime.date.today()}{Style.RESET_ALL}")
    print(f"Reviewing {len(props)} propert{'y' if len(props)==1 else 'ies'}...")
    print(f"{Fore.YELLOW}Note: ~3 min per property (Playwright renders each page fully).{Style.RESET_ALL}\n")

    results = []
    for prop in props:
        try:
            results.append(review_property(prop))
            print_result(results[-1])
        except Exception as exc:
            print(f"{Fore.RED}ERROR reviewing {prop['name']}: {exc}{Style.RESET_ALL}")

    if not args.no_excel and results:
        write_excel(results, f"report_{datetime.date.today().isoformat()}.xlsx")

    print(f"\n{Fore.WHITE}{Style.BRIGHT}Review complete.{Style.RESET_ALL}\n")


if __name__ == "__main__":
    main()
