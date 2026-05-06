"""
Defense Procurement + GeM Tender Scraper
==========================================
Portals : https://defproc.gov.in/nicgep/app   (DefProc / NIC GePNIC)
          https://gem.gov.in/                  (Government e-Marketplace)
Method  : Selenium (headless Chrome) — both portals scraped in PARALLEL
Output  : tenders_output.xlsx

Install:
    pip install selenium webdriver-manager openpyxl beautifulsoup4

Run:
    python aerotender.py
"""

import os
import stat
import time
import re as _re
import concurrent.futures
from datetime import datetime
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── KEYWORD HIGHLIGHT HELPER ──────────────────────────────────────────────────

def contains_keyword(text: str, keywords: list) -> str:
    """
    Returns the first matched keyword (case-insensitive) found in text,
    or empty string if none match.
    """
    for kw in sorted(keywords, key=len, reverse=True):
        if _re.search(_re.escape(kw), text, _re.IGNORECASE):
            return kw
    return ""

# ── KEYWORDS ──────────────────────────────────────────────────────────────────

KEYWORDS = [
    "aero engine", "aeroengine", "thrust", "turbojet", "turbofan",
    "turboshaft", "forge", "forging", "casting", "foundry",
    "ramjet", "scramjet", "superjet", "amca", "amagb",
    "accessory gear box", "hindustan aeronautics limited",
    "gas turbine", "power plant",
]

DEFPROC_BASE_URL = "https://defproc.gov.in/nicgep/app"
GEM_BASE_URL     = "https://gem.gov.in/"
GEM_SEARCH_URL   = "https://bidplus.gem.gov.in/search"
OUTPUT_FILE      = "tenders_output.xlsx"

# ── MONTH HELPERS  (NEW) ──────────────────────────────────────────────────────

_MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3,  "apr": 4,
    "may": 5, "jun": 6, "jul": 7,  "aug": 8,
    "sep": 9, "oct":10, "nov":11,  "dec":12,
}
_MONTH_NAME = {v: k.capitalize() for k, v in _MONTH_MAP.items()}

def _current_month_range_label() -> str:
    """
    Returns a human-readable filter label, e.g. 'May–Dec 2026'.
    Recalculated every run so it always reflects today's date.
    """
    now = datetime.now()
    return f"{_MONTH_NAME[now.month]}–Dec {now.year}"

# ── CLOSING DATE FILTER  (UPDATED) ───────────────────────────────────────────
#
#  *** CHANGE from previous version ***
#  Previously hard-coded to Apr/May/Jun only.
#  Now keeps any tender whose Due Date falls between the CURRENT calendar month
#  and December of the CURRENT year (inclusive).
#  Tenders closing in a FUTURE year are also kept.
#
#  Supported date formats:
#    DD-Mon-YYYY HH:MM AM/PM  →  '18-Jun-2026 10:00 AM'
#    DD-Mon-YYYY              →  '18-Jun-2026'
#    DD/Mon/YYYY              →  '18/Jun/2026'

def is_may_or_june(due_date_str: str) -> bool:
    """
    Returns True when the closing date is within the window:
        current month  ≤  due date  ≤  December of current year
    OR the due date falls in any future year.
    """
    if not due_date_str:
        return False
    try:
        normalised = due_date_str.strip().replace("/", "-")
        parts      = normalised.split("-")          # ['18', 'Jun', '2026 10:00 AM']
        month_str  = parts[1][:3].lower()
        year       = int(parts[2][:4])
        month_num  = _MONTH_MAP.get(month_str, 0)

        if month_num == 0:
            return False

        now           = datetime.now()
        current_month = now.month
        current_year  = now.year

        # Future year → always include
        if year > current_year:
            return True
        # Same year → include only from current month onward (up to Dec)
        if year == current_year and month_num >= current_month:
            return True
        return False

    except Exception:
        return False

# ── DRIVER SETUP ──────────────────────────────────────────────────────────────
# *** FIX: resolve the driver path first, explicitly grant execute permission
#          (chmod 755) before handing it to the Service constructor.
#          This resolves the "wrong permissions" crash on Windows & Linux alike.

def create_driver():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )

    # ── Railway / Docker: prefer system Chrome set via env vars ─────────────
    chrome_bin       = os.environ.get("CHROME_BIN")
    chromedriver_path = os.environ.get("CHROMEDRIVER_PATH")

    if chrome_bin:
        options.binary_location = chrome_bin

    if chromedriver_path:
        # System chromedriver (e.g. /usr/bin/chromedriver on Railway)
        driver_path = chromedriver_path
        print(f"  [INFO] Using system chromedriver: {driver_path}")
    else:
        # Local dev: auto-download via webdriver-manager
        driver_path = ChromeDriverManager().install()
        try:
            os.chmod(
                driver_path,
                stat.S_IRWXU | stat.S_IRGRP | stat.S_IXGRP |
                stat.S_IROTH | stat.S_IXOTH,   # 755
            )
        except Exception as perm_err:
            print(f"  [WARN] chmod on chromedriver skipped: {perm_err}")

    driver = webdriver.Chrome(
        service=Service(driver_path), options=options
    )
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"}
    )
    return driver

# ── FIND ELEMENT WITH MULTIPLE FALLBACK SELECTORS ────────────────────────────

def find_element_any(driver, selectors_css, timeout=10):
    """Try each CSS selector in order, return first match found."""
    end = time.time() + timeout
    while time.time() < end:
        for css in selectors_css:
            try:
                el = driver.find_element(By.CSS_SELECTOR, css)
                if el.is_displayed():
                    return el
            except Exception:
                pass
        time.sleep(0.5)
    return None

# ── SEARCH INPUT SELECTORS — covers all common NIC GePNIC variants ────────────

SEARCH_INPUT_SELECTORS = [
    "input[name='keyword']",
    "input[id='keyword']",
    "input[name='searchText']",
    "input[id='searchText']",
    "input[name='tenderTitle']",
    "input[id='tenderTitle']",
    "input[name='productCategory']",
    "input[placeholder*='keyword' i]",
    "input[placeholder*='search' i]",
    "input[placeholder*='tender' i]",
    "input[type='text']",
]

GO_BUTTON_SELECTORS = [
    "input[value='Go']",
    "input[value='GO']",
    "input[value='go']",
    "button[type='submit']",
    "input[type='submit']",
    "a[href*='search']",
    "button:not([type='reset'])",
]

# ══════════════════════════════════════════════════════════════════════════════
#  DEFPROC PORTAL SCRAPER  (original — unchanged)
# ══════════════════════════════════════════════════════════════════════════════

def search_keyword(driver, keyword):
    results = []
    try:
        driver.get(DEFPROC_BASE_URL)
        time.sleep(2)

        search_input = find_element_any(driver, SEARCH_INPUT_SELECTORS, timeout=10)
        if search_input is None:
            all_inputs = driver.find_elements(By.TAG_NAME, "input")
            print(f"\n    [DEBUG] Inputs on page:")
            for inp in all_inputs:
                print(f"      type={inp.get_attribute('type')}  "
                      f"name={inp.get_attribute('name')}  "
                      f"id={inp.get_attribute('id')}  "
                      f"placeholder={inp.get_attribute('placeholder')}")
            raise Exception("Search input not found — see debug output above")

        search_input.click()
        time.sleep(0.3)
        search_input.clear()
        search_input.send_keys(keyword)
        time.sleep(0.3)

        go_btn = find_element_any(driver, GO_BUTTON_SELECTORS, timeout=5)
        if go_btn:
            go_btn.click()
        else:
            search_input.send_keys(Keys.RETURN)

        time.sleep(3)

        page = 1
        while True:
            soup = BeautifulSoup(driver.page_source, "html.parser")
            page_results = parse_defproc_table(soup, keyword)
            results.extend(page_results)

            next_clicked = False
            next_selectors = [
                "a[title='Next']", "a.next", "a[rel='next']",
                "//a[contains(text(),'Next')]",
                "//a[contains(text(),'>')]",
            ]
            for sel in next_selectors:
                try:
                    if sel.startswith("//"):
                        el = driver.find_element(By.XPATH, sel)
                    else:
                        el = driver.find_element(By.CSS_SELECTOR, sel)
                    if el.is_displayed() and el.is_enabled():
                        el.click()
                        time.sleep(2)
                        page += 1
                        next_clicked = True
                        break
                except Exception:
                    pass

            if not next_clicked:
                break

    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"\n    [WARN] DefProc '{keyword}': {e}")

    return results


_TENDER_RE = _re.compile(
    r'\d+\.\s+'
    r'(\d{2}-\w{3}-\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)\s+'
    r'(\d{2}-\w{3}-\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)\s+'
    r'(\d{2}-\w{3}-\d{4}\s+\d{1,2}:\d{2}\s+[AP]M)\s+'
    r'\[([^\]]+)\]\s+'
    r'\[([^\]]+)\]'
    r'\[([^\]]+)\]\s*'
    r'(.+?)(?=\s*\d+\.\s+\d{2}-\w{3}-\d{4}|\Z)',
    _re.DOTALL
)


def parse_defproc_table(soup, keyword):
    results = []
    page_link = DEFPROC_BASE_URL
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "FrontEndAdvancedSearchResult" in href or "searchResult" in href.lower():
            page_link = href if href.startswith("http") else "https://defproc.gov.in" + href
            break

    full_text = soup.get_text(separator=" ")

    for m in _TENDER_RE.finditer(full_text):
        pub_date   = m.group(1).strip()
        close_date = m.group(2).strip()
        title      = m.group(4).strip()
        ref_no     = m.group(5).strip()
        tender_id  = m.group(6).strip()
        org_chain  = m.group(7).strip()
        org        = org_chain.split("||")[0].strip()

        if tender_id:
            link = (
                f"https://defproc.gov.in/nicgep/app"
                f"?component=%24DirectLink"
                f"&page=FrontEndTenderDetailswithoutLogin"
                f"&service=direct&session=T&sp={tender_id}"
            )
        else:
            link = page_link

        if not title:
            continue

        results.append({
            "Source":                    "DefProc",
            "Description of the Tender": title,
            "Tender Floated By":         org,
            "Ref.No":                    ref_no,
            "Tender Date":               pub_date,
            "Due Date":                  close_date,
            "Portal/Web link":           link,
            "_keyword":                  keyword,
        })

    return results


def scrape_defproc_all():
    """Scrape the DefProc portal for all keywords."""
    driver      = create_driver()
    all_tenders = []
    seen        = set()
    date_label  = _current_month_range_label()

    print(f"\n{'='*55}")
    print(f"  [DefProc] Portal    : {DEFPROC_BASE_URL}")
    print(f"  [DefProc] Keywords  : {len(KEYWORDS)}")
    print(f"  [DefProc] Date range: {date_label}")
    print(f"{'='*55}\n")

    try:
        for i, kw in enumerate(KEYWORDS, 1):
            print(f"  [DefProc {i:02d}/{len(KEYWORDS)}] '{kw}' ...", end=" ", flush=True)
            rows = search_keyword(driver, kw)

            added = 0
            for r in rows:
                if not is_may_or_june(r.get("Due Date", "")):
                    continue
                key = (r.get("Ref.No") or r.get("Description of the Tender","")[:60]).strip()
                if key and key not in seen:
                    seen.add(key)
                    all_tenders.append(r)
                    added += 1

            print(f"  {len(rows)} scraped, {added} match {date_label} closing")

    except KeyboardInterrupt:
        print(f"\n  [!] DefProc interrupted — saving {len(all_tenders)} tenders so far ...")
    finally:
        driver.quit()

    print(f"\n  [DefProc] Total unique tenders: {len(all_tenders)}\n")
    return all_tenders


# ══════════════════════════════════════════════════════════════════════════════
#  GEM PORTAL SCRAPER
# ══════════════════════════════════════════════════════════════════════════════

def scrape_gem_keyword(driver, keyword):
    """Search GeM bidplus portal for one keyword, return list of tender dicts."""
    results = []
    try:
        url = f"{GEM_SEARCH_URL}?searchedText={keyword.replace(' ', '+')}"
        driver.get(url)
        time.sleep(4)

        page = 1
        while True:
            soup = BeautifulSoup(driver.page_source, "html.parser")
            page_results = parse_gem_page(soup, keyword)
            results.extend(page_results)

            next_clicked = False
            next_selectors = [
                "a.next", "li.next > a", "a[aria-label='Next']",
                "//a[contains(text(),'Next')]",
                "//button[contains(text(),'Next')]",
                "//li[contains(@class,'next')]/a",
            ]
            for sel in next_selectors:
                try:
                    el = (driver.find_element(By.XPATH, sel)
                          if sel.startswith("//")
                          else driver.find_element(By.CSS_SELECTOR, sel))
                    if el.is_displayed() and el.is_enabled():
                        driver.execute_script("arguments[0].click();", el)
                        time.sleep(2)
                        page += 1
                        next_clicked = True
                        break
                except Exception:
                    pass

            if not next_clicked:
                break

    except KeyboardInterrupt:
        raise
    except Exception as e:
        print(f"\n    [WARN] GeM '{keyword}': {e}")

    return results


def parse_gem_page(soup, keyword):
    """
    Parse GeM bid search results page.
    GeM renders bid cards inside <div class='bidding-list-card'> or similar.
    Multiple fallback strategies are attempted to handle site updates.
    """
    results = []

    card_selectors = [
        lambda tag: tag.name == "div" and tag.get("class") and
                    any("card" in c.lower() or "bid" in c.lower() for c in tag.get("class", [])),
        lambda tag: tag.name == "li" and tag.get("class") and
                    any("item" in c.lower() or "result" in c.lower() for c in tag.get("class", [])),
    ]

    cards = []
    for sel in card_selectors:
        cards = soup.find_all(sel)
        if cards:
            break

    for card in cards:
        try:
            raw_text = card.get_text(separator=" ", strip=True)

            title = ""
            for tag in ["h3", "h4", "h5", "h2"]:
                el = card.find(tag)
                if el:
                    title = el.get_text(strip=True)
                    break
            if not title:
                el = card.find(["strong", "b"])
                if el:
                    title = el.get_text(strip=True)

            bid_no = ""
            bid_match = _re.search(r'GEM[-/]\d{4}[-/][A-Z0-9]+[-/][A-Z0-9]+', raw_text)
            if bid_match:
                bid_no = bid_match.group(0)
            else:
                bid_match2 = _re.search(r'BID[/\-][A-Z0-9/\-]{6,}', raw_text)
                if bid_match2:
                    bid_no = bid_match2.group(0)

            org = ""
            for klass_hint in ["ministry", "org", "department", "buyer"]:
                el = card.find(class_=_re.compile(klass_hint, _re.I))
                if el:
                    org = el.get_text(strip=True)
                    break

            date_matches = _re.findall(r'\d{2}[-/]\w{3}[-/]\d{4}', raw_text)
            tender_date  = date_matches[0] if len(date_matches) > 0 else ""
            due_date_str = date_matches[-1] if len(date_matches) > 1 else ""

            link = ""
            a_tag = card.find("a", href=True)
            if a_tag:
                href = a_tag["href"]
                if href.startswith("http"):
                    link = href
                elif href.startswith("/"):
                    link = "https://bidplus.gem.gov.in" + href
            if not link and bid_no:
                safe_bid = bid_no.replace("/", "%2F")
                link = f"https://bidplus.gem.gov.in/showbiddetail/{safe_bid}"

            if not title and not bid_no:
                continue

            results.append({
                "Source":                    "GeM",
                "Description of the Tender": title or bid_no,
                "Tender Floated By":         org,
                "Ref.No":                    bid_no,
                "Tender Date":               tender_date,
                "Due Date":                  due_date_str,
                "Portal/Web link":           link,
                "_keyword":                  keyword,
            })
        except Exception:
            continue

    if not results:
        full_text = soup.get_text(separator=" ")
        for m in _re.finditer(
            r'(GEM[-/]\d{4}[-/][A-Z]\d*/\d+)'
            r'(.{0,300}?)'
            r'(\d{2}[-/]\w{3}[-/]\d{4})',
            full_text, _re.DOTALL
        ):
            bid_no     = m.group(1).strip()
            context    = m.group(2).strip()
            due_date_s = m.group(3).strip()
            safe_bid   = bid_no.replace("/", "%2F")
            results.append({
                "Source":                    "GeM",
                "Description of the Tender": context[:120] if context else bid_no,
                "Tender Floated By":         "",
                "Ref.No":                    bid_no,
                "Tender Date":               "",
                "Due Date":                  due_date_s,
                "Portal/Web link":           f"https://bidplus.gem.gov.in/showbiddetail/{safe_bid}",
                "_keyword":                  keyword,
            })

    return results


def scrape_gem_all():
    """Scrape the GeM portal for all keywords."""
    driver      = create_driver()
    all_tenders = []
    seen        = set()
    date_label  = _current_month_range_label()

    print(f"\n{'='*55}")
    print(f"  [GeM] Portal    : {GEM_BASE_URL}")
    print(f"  [GeM] Keywords  : {len(KEYWORDS)}")
    print(f"  [GeM] Date range: {date_label}")
    print(f"{'='*55}\n")

    try:
        for i, kw in enumerate(KEYWORDS, 1):
            print(f"  [GeM {i:02d}/{len(KEYWORDS)}] '{kw}' ...", end=" ", flush=True)
            rows = scrape_gem_keyword(driver, kw)

            added = 0
            for r in rows:
                if not is_may_or_june(r.get("Due Date", "")):
                    continue
                key = (r.get("Ref.No") or r.get("Description of the Tender", "")[:60]).strip()
                if key and key not in seen:
                    seen.add(key)
                    all_tenders.append(r)
                    added += 1

            print(f"  {len(rows)} scraped, {added} match {date_label} closing")

    except KeyboardInterrupt:
        print(f"\n  [!] GeM interrupted — saving {len(all_tenders)} tenders so far ...")
    finally:
        driver.quit()

    print(f"\n  [GeM] Total unique tenders: {len(all_tenders)}\n")
    return all_tenders


# ══════════════════════════════════════════════════════════════════════════════
#  PARALLEL ORCHESTRATOR
# ══════════════════════════════════════════════════════════════════════════════

def scrape_all():
    """
    Launches DefProc and GeM scrapers in parallel threads.
    Each thread creates its own Chrome driver instance.
    Returns combined list of all unique tenders.
    """
    date_label = _current_month_range_label()
    print(f"\n  Launching BOTH portals in parallel ...")
    print(f"  Date filter  : closing date between {date_label}\n")

    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        future_defproc = executor.submit(scrape_defproc_all)
        future_gem     = executor.submit(scrape_gem_all)

        try:
            defproc_tenders = future_defproc.result()
        except Exception as e:
            print(f"  [ERROR] DefProc scraper failed: {e}")
            defproc_tenders = []

        try:
            gem_tenders = future_gem.result()
        except Exception as e:
            print(f"  [ERROR] GeM scraper failed: {e}")
            gem_tenders = []

    combined = defproc_tenders + gem_tenders

    print(f"\n{'='*55}")
    print(f"  DefProc tenders ({date_label}) : {len(defproc_tenders)}")
    print(f"  GeM tenders     ({date_label}) : {len(gem_tenders)}")
    print(f"  COMBINED TOTAL                 : {len(combined)}")
    print(f"{'='*55}\n")

    return combined


# ── EXCEL EXPORT ──────────────────────────────────────────────────────────────

COLUMNS = [
    "Source",
    "Description of the Tender",
    "Tender Floated By",
    "Ref.No",
    "Tender Date",
    "Due Date",
    "Portal/Web link",
]

COL_W = {
    "Source":                        10,
    "Description of the Tender":     52,
    "Tender Floated By":             30,
    "Ref.No":                        22,
    "Tender Date":                   16,
    "Due Date":                      16,
    "Portal/Web link":               20,
}

SOURCE_COLORS = {
    "DefProc": ("1F3864", "FFFFFF"),
    "GeM":     ("1D6F42", "FFFFFF"),
}

THIN   = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_excel(data):
    wb         = Workbook()
    ws         = wb.active
    ws.title   = "Tenders"
    date_label = _current_month_range_label()
    num_cols   = len(COLUMNS)
    last_col   = get_column_letter(num_cols)

    # ── Title row ─────────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    t = ws["A1"]
    t.value = (
        f"Defense Procurement Tenders  |  Aero-Engine / Aerospace Keywords  |  "
        f"Closing: {date_label}  |  "
        f"Sources: DefProc + GeM  |  "
        f"{datetime.now().strftime('%d-%b-%Y  %H:%M')}"
    )
    t.font      = Font(name="Arial", bold=True, size=12, color="FFFFFF")
    t.fill      = PatternFill("solid", start_color="0D2137", end_color="0D2137")
    t.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 34

    # ── Header row ────────────────────────────────────────────────────────────
    hdr_fill = PatternFill("solid", start_color="1F3864", end_color="1F3864")
    for ci, col in enumerate(COLUMNS, 1):
        c = ws.cell(row=2, column=ci, value=col)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill      = hdr_fill
        c.border    = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 24

    alt_fill = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")

    # ── Data rows ─────────────────────────────────────────────────────────────
    for ri, rec in enumerate(data, start=3):
        alt    = (ri % 2 == 0)
        source = rec.get("Source", "")

        for ci, col in enumerate(COLUMNS, 1):
            val  = rec.get(col, "")
            cell = ws.cell(row=ri, column=ci)
            cell.border    = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            base_font = Font(name="Arial", size=10)

            if col == "Source":
                bg, fg = SOURCE_COLORS.get(source, ("888888", "FFFFFF"))
                cell.value     = source
                cell.fill      = PatternFill("solid", start_color=bg, end_color=bg)
                cell.font      = Font(name="Arial", size=10, bold=True, color=fg)
                cell.alignment = Alignment(horizontal="center", vertical="top")

            elif col == "Description of the Tender" and val:
                cell.value = val
                matched = contains_keyword(val, KEYWORDS)
                if matched:
                    cell.font = Font(name="Arial", size=10, bold=True, color="CC0000")
                else:
                    cell.font = base_font
                if alt:
                    cell.fill = alt_fill

            elif col == "Portal/Web link" and str(val).startswith("http"):
                cell.value     = "Open Tender ↗"
                cell.hyperlink = val
                cell.font = Font(
                    name="Arial", size=10,
                    color="0563C1", underline="single", bold=True
                )
                cell.alignment = Alignment(horizontal="center", vertical="top")
                if alt:
                    cell.fill = alt_fill

            else:
                cell.value = val
                cell.font  = base_font
                if alt:
                    cell.fill = alt_fill

    for ci, col in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = COL_W[col]

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{last_col}2"

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Summary")

    for ci, h in enumerate(["Source", "Keyword", "Tenders Found"], 1):
        c = ws2.cell(row=1, column=ci, value=h)
        c.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill      = hdr_fill
        c.alignment = Alignment(horizontal="center")

    kw_count = {}
    for rec in data:
        key = (rec.get("Source", "?"), rec.get("_keyword", ""))
        kw_count[key] = kw_count.get(key, 0) + 1

    for ri, ((src, kw), cnt) in enumerate(sorted(kw_count.items()), start=2):
        ws2.cell(row=ri, column=1, value=src).font  = Font(name="Arial", size=10)
        ws2.cell(row=ri, column=2, value=kw).font   = Font(name="Arial", size=10)
        ws2.cell(row=ri, column=3, value=cnt).font  = Font(name="Arial", size=10)

    tr = len(kw_count) + 2
    ws2.cell(row=tr, column=1, value="TOTAL").font   = Font(name="Arial", bold=True, size=10)
    ws2.cell(row=tr, column=3, value=len(data)).font = Font(name="Arial", bold=True, size=10)

    ws2.column_dimensions["A"].width = 12
    ws2.column_dimensions["B"].width = 42
    ws2.column_dimensions["C"].width = 16

    wb.save(OUTPUT_FILE)
    print(f"  Excel saved -> {OUTPUT_FILE}  ({len(data)} tenders total)\n")


# ── ENTRY POINT ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        tenders = scrape_all()
        if tenders:
            write_excel(tenders)
        else:
            print("No tenders found matching the date filter.")
            print("Check [DEBUG] output above for input field details.")
            print("Update SEARCH_INPUT_SELECTORS if the portal layout changed.")
    except KeyboardInterrupt:
        print("\n  [!] Aborted before any data was collected. No file saved.")
