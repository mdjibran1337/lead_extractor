import re
import asyncio
import aiohttp
import time
import os
import sys
import traceback
from urllib.parse import urlparse
from fastapi import FastAPI, Form
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

app = FastAPI(title="Lead Engine")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"],
)

# ---------------- CONFIG ----------------
EMAIL_REGEX = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
PHONE_REGEX = r'(?:\+91[\-\s]?)?[6-9]\d{9}|0\d{2,4}[\-\s]?\d{6,8}'

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'}
CONTACT_PATHS = ['/contact', '/contact-us', '/about', '/about-us', '/reach-us']

INVALID_EMAIL_SUFFIX = ('.png', '.jpg', '.jpeg', '.svg', '.gif', '.webp', '.css', '.js', '.ico', '.bmp', '.woff', '.woff2', '.ttf', '.eot')

BLOCKED_DOMAINS = (
    'google.com', 'google.co.in', 'google.co', 'googleapis.com',
    'gstatic.com', 'gmail.com', 'youtube.com',
    'support.google', 'policies.google', 'about.google',
    'accounts.google', 'myaccount.google',
    '/maps/place', '/maps/search', '/intl/en/about',
    'facebook.com/sharer', 'twitter.com/intent',
    'instagram.com/accounts', 'linkedin.com/sharing',
)

EXCEL_FILE = "ultimate_master_leads.xlsx"

# YAHAN CHANGE KIYA HAI: Default "0" ki jagah "1" kar diya hai taaki server pe headless chale
HEADLESS = os.environ.get("HEADLESS", "1") == "1"


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


# ---------------- VALIDATORS ----------------
def is_valid_email(e):
    if not e or '@' not in e:
        return False
    low = e.lower()
    if low.endswith(INVALID_EMAIL_SUFFIX):
        return False
    if any(x in low for x in ('example.com', 'domain.com', 'sentry.io', 'wixpress.com', 'yourdomain')):
        return False
    return True


def is_valid_phone(raw):
    if not raw:
        return False
    digits = re.sub(r'\D', '', str(raw))
    if len(digits) == 10 and digits[0] in '6789':
        return True
    if len(digits) == 12 and digits.startswith('91') and digits[2] in '6789':
        return True
    if len(digits) == 11 and digits.startswith('0'):
        return True
    return False


def normalize_phone(raw):
    if not is_valid_phone(raw):
        return None
    digits = re.sub(r'\D', '', str(raw))
    if len(digits) == 10 and digits[0] in '6789':
        return f"+91 {digits[:5]} {digits[5:]}"
    if len(digits) == 12 and digits.startswith('91') and digits[2] in '6789':
        return f"+91 {digits[2:7]} {digits[7:]}"
    if len(digits) == 11 and digits.startswith('0'):
        mobile = digits[1:]
        if mobile[0] in '6789':
            return f"+91 {mobile[:5]} {mobile[5:]}"
        return f"0{digits[1:5]} {digits[5:]}"
    return digits


def is_valid_website(href):
    if not href:
        return False
    h = href.lower()
    if not h.startswith('http'):
        return False
    for b in BLOCKED_DOMAINS:
        if b in h:
            return False
    return True


def get_clean_domain_name(url):
    try:
        p = urlparse(url if url.startswith(('http://', 'https://')) else 'https://' + url)
        host = p.netloc if p.netloc else p.path
        host = host.replace('www.', '').split('.')[0]
        return host.capitalize() if host else 'Website'
    except Exception:
        return 'Website'


def save_excel(results):
    try:
        df = pd.DataFrame(results)
        cols = ["Business Name", "Phone Number", "Website", "Official Email ID",
                "Address", "Rating", "Source"]
        df = df[[c for c in cols if c in df.columns]]
        df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        log(f"Excel save error: {e}")


def error_row(msg, url="N/A", source="Error"):
    return [{"Business Name": get_clean_domain_name(url) if url != "N/A" else "Error",
             "Phone Number": "N/A", "Website": url, "Official Email ID": "N/A",
             "Address": msg, "Rating": "N/A", "Source": source}]


# ---------------- SELENIUM BASE ----------------
def launch_driver():
    try:
        opts = webdriver.ChromeOptions()
        opts.add_argument('--disable-blink-features=AutomationControlled')
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option('useAutomationExtension', False)
        opts.add_argument('--no-sandbox')
        opts.add_argument('--disable-dev-shm-usage')
        opts.add_argument('--disable-gpu')
        opts.add_argument('--log-level=3')
        opts.add_argument(f'user-agent={HEADERS["User-Agent"]}')
        if HEADLESS:
            opts.add_argument('--headless=new')
        # Selenium 4.6+ automatically handles the driver. No webdriver_manager needed.
        return webdriver.Chrome(options=opts), None
    except Exception as e:
        return None, f"Launch fail: {e}"


def find_first(driver, selectors, timeout=3, want_many=False):
    end = time.time() + timeout
    while time.time() < end:
        for sel in selectors:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                if els:
                    return els if want_many else els[0]
            except Exception:
                pass
        time.sleep(0.3)
    return [] if want_many else None


def handle_consent(driver):
    for xp in ["//button[contains(., 'Accept all')]",
               "//button[contains(., 'Reject all')]",
               "//button[contains(., 'I agree')]",
               "//button[contains(., 'Accept')]"]:
        try:
            driver.find_element(By.XPATH, xp).click()
            time.sleep(1)
            return
        except Exception:
            continue


# ---------------- GMAPS: SINGLE PLACE EXTRACTOR ----------------
def extract_place_details(driver):
    name_el = find_first(driver, ['h1.DUwDvf', 'h1[class*="fontHeadlineLarge"]', 'h1'], timeout=5)
    name = name_el.text.strip() if name_el else ""
    if not name:
        return None

    phone_val = "N/A"
    phone_btn = find_first(driver, [
        'button[data-item-id^="phone"]',
        'button[data-tooltip*="phone" i]',
        'a[href^="tel:"]',
    ], timeout=2)
    if phone_btn:
        raw = phone_btn.text or phone_btn.get_attribute('aria-label') or ""
        raw = re.sub(r'^\s*phone\s*:\s*', '', raw, flags=re.I).strip()
        m = re.search(r'(\+?\d[\d\s\-\(\)]{8,}\d)', raw)
        if m:
            formatted = normalize_phone(m.group(1))
            if formatted:
                phone_val = formatted

    website_val = "N/A"
    web_el = find_first(driver, [
        'a[data-item-id="authority"]',
        'a[data-item-id*="authority"]',
        'a[aria-label*="website" i][href^="http"]',
    ], timeout=2)
    if web_el:
        href = web_el.get_attribute('href') or ""
        href = href.split('?')[0].rstrip('/')
        if is_valid_website(href):
            website_val = href

    address_val = "N/A"
    addr_el = find_first(driver, [
        'button[data-item-id="address"]',
        'button[data-tooltip*="address" i]',
    ], timeout=2)
    if addr_el:
        txt = addr_el.text or addr_el.get_attribute('aria-label') or ""
        txt = re.sub(r'^\s*address\s*:\s*', '', txt, flags=re.I)
        txt = re.sub(r'\s+', ' ', txt).strip()
        parts = [p.strip() for p in txt.split('·') if p.strip()]
        if parts:
            address_val = parts[-1] if len(parts) > 1 else parts[0]

    rating_val = "N/A"
    r_el = find_first(driver, ['div.F7nice span[aria-hidden="true"]', 'span[aria-label*="stars" i]'], timeout=2)
    if r_el:
        t = r_el.text.strip()
        if re.match(r'^[1-5](\.\d)?$', t):
            rating_val = t

    return {
        "Business Name": name,
        "Phone Number": phone_val,
        "Website": website_val,
        "Official Email ID": "N/A",
        "Address": address_val,
        "Rating": rating_val,
        "Source": "Google Maps Live"
    }


def run_selenium_gmaps(search_query, max_results=10):
    log(f"=== GMAPS START: {search_query} ===")
    driver, err = launch_driver()
    if not driver:
        return error_row(f"Launch failed: {err[:200]}", source="Gmaps — launch error")

    results = []
    seen_names = set()
    try:
        url = f"https://www.google.com/maps/search/{search_query.replace(' ', '+')}"
        log(f"Opening: {url}")
        driver.get(url)
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="feed"], div[role="main"], h1'))
            )
        except Exception:
            pass
        handle_consent(driver)
        time.sleep(2)

        feed = find_first(driver, ['div[role="feed"]'], timeout=5)
        if feed:
            for _ in range(5):
                try:
                    driver.execute_script('arguments[0].scrollTop = arguments[0].scrollHeight', feed)
                    time.sleep(2)
                except Exception:
                    break

        links = driver.find_elements(By.CSS_SELECTOR, 'a[href*="/maps/place/"]')
        if not links:
            links = driver.find_elements(By.CSS_SELECTOR, 'a.hf2wfc')

        place_urls = []
        for l in links:
            href = l.get_attribute('href')
            if href and '/maps/place/' in href and href not in place_urls:
                place_urls.append(href)

        log(f"Found {len(place_urls)} unique place URLs")

        for idx, place_url in enumerate(place_urls[:max_results * 2]):
            if len(results) >= max_results:
                break
            try:
                driver.get(place_url)
                time.sleep(2.5)
                detail = extract_place_details(driver)
                if not detail:
                    log(f"  [{idx}] no name, skip")
                    continue
                if detail["Business Name"] in seen_names:
                    log(f"  [{idx}] dup skip: {detail['Business Name']}")
                    continue
                seen_names.add(detail["Business Name"])
                results.append(detail)
                log(f"  [{idx}] {detail['Business Name']} | phone={detail['Phone Number']} | "
                    f"site={detail['Website'][:50] if detail['Website'] != 'N/A' else 'N/A'}")
            except Exception as e:
                log(f"  [{idx}] error: {e}")
                continue
    except Exception as e:
        log(f"GMAPS FATAL: {e}\n{traceback.format_exc()[:300]}")
        return error_row(f"Gmaps error: {e}", source="Gmaps — exception")
    finally:
        try:
            driver.quit()
        except Exception:
            pass

    log(f"=== GMAPS DONE: {len(results)} unique businesses ===")
    return results


# ---------------- WEB SCRAPER (Selenium) ----------------
def _wait_dom_stable(driver, timeout=12, stable_rounds=3):
    end = time.time() + timeout
    last, cnt = -1, 0
    while time.time() < end:
        try:
            size = driver.execute_script("return document.body ? document.body.scrollHeight : 0")
        except Exception:
            size = 0
        if size == last:
            cnt += 1
            if cnt >= stable_rounds:
                return True
        else:
            cnt, last = 0, size
        time.sleep(0.7)
    return False


def _auto_scroll(driver, passes=6, pause=1.2):
    for _ in range(passes):
        try:
            old = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(pause)
            new = driver.execute_script("return document.body.scrollHeight")
            if new == old:
                break
        except Exception:
            break


def _click_expand_buttons(driver):
    xp = ("//button["
          "contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show more') or "
          "contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'load more') or "
          "contains(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'show phone')]")
    try:
        for b in driver.find_elements(By.XPATH, xp)[:10]:
            try:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", b)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", b)
                time.sleep(1)
            except Exception:
                continue
    except Exception:
        pass


def _extract_contacts(driver):
    emails, phones = set(), set()
    try:
        for a in driver.find_elements(By.CSS_SELECTOR, 'a[href^="mailto:"]'):
            e = (a.get_attribute('href') or '').replace('mailto:', '').split('?')[0].strip().lower()
            if is_valid_email(e):
                emails.add(e)
    except Exception:
        pass
    try:
        for a in driver.find_elements(By.CSS_SELECTOR, 'a[href^="tel:"]'):
            raw = (a.get_attribute('href') or '').replace('tel:', '').strip()
            p = normalize_phone(raw)
            if p:
                phones.add(p)
    except Exception:
        pass
    for sel in ['script[type="application/ld+json"]', '#__NEXT_DATA__', '#__NUXT_DATA__']:
        try:
            for sc in driver.find_elements(By.CSS_SELECTOR, sel):
                txt = sc.get_attribute('innerHTML') or ''
                for e in re.findall(EMAIL_REGEX, txt):
                    if is_valid_email(e):
                        emails.add(e.lower())
                for p in re.findall(PHONE_REGEX, txt):
                    c = normalize_phone(p)
                    if c:
                        phones.add(c)
        except Exception:
            pass
    try:
        for e in re.findall(EMAIL_REGEX, driver.page_source):
            if is_valid_email(e):
                emails.add(e.lower())
    except Exception:
        pass
    try:
        body = driver.find_element(By.TAG_NAME, 'body').text
        for p in re.findall(PHONE_REGEX, body):
            c = normalize_phone(p)
            if c:
                phones.add(c)
    except Exception:
        pass
    return emails, phones


def run_selenium_web_scraper(target_url, crawl_contact=True):
    log(f"=== WEB START: {target_url} ===")
    if not target_url.startswith(('http://', 'https://')):
        target_url = 'https://' + target_url
    driver, err = launch_driver()
    if not driver:
        return error_row(f"Launch failed: {err[:200]}", target_url, source="Web — launch error")

    results = []
    try:
        driver.get(target_url)
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        except Exception:
            pass
        _wait_dom_stable(driver, timeout=12)
        _auto_scroll(driver, passes=6, pause=1)
        _click_expand_buttons(driver)
        time.sleep(1)

        emails, phones = _extract_contacts(driver)
        try:
            body_len = len(driver.find_element(By.TAG_NAME, 'body').text)
        except Exception:
            body_len = 0
        log(f"WEB title='{driver.title[:60]}' body_len={body_len} emails={len(emails)} phones={len(phones)}")

        if crawl_contact:
            p = urlparse(target_url)
            base = f"{p.scheme}://{p.netloc}"
            for path in CONTACT_PATHS:
                try:
                    driver.get(base + path)
                    _wait_dom_stable(driver, timeout=6)
                    time.sleep(0.5)
                    e2, ph2 = _extract_contacts(driver)
                    emails |= e2
                    phones |= ph2
                except Exception:
                    continue

        biz = get_clean_domain_name(target_url)
        if not emails and not phones:
            results = error_row(
                f"No contacts found. body_len={body_len}, title='{driver.title[:50]}'. "
                f"Site likely has no public emails/phones.",
                target_url, source="Web — empty"
            )
        else:
            results.append({
                "Business Name": biz,
                "Phone Number": ", ".join(sorted(phones)) if phones else "N/A",
                "Website": target_url,
                "Official Email ID": ", ".join(sorted(emails)) if emails else "N/A",
                "Address": "N/A", "Rating": "N/A",
                "Source": "Selenium Web Scraper"
            })
    except Exception as e:
        log(f"WEB ERROR: {e}")
        results = error_row(f"Error: {str(e)[:200]}", target_url, source="Web — exception")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
    log(f"=== WEB DONE: {len(results)} rows ===")
    return results


# ---------------- ASYNC SCRAPER ----------------
async def fetch_page(session, url, timeout=12):
    try:
        if not url.startswith(('http://', 'https://')):
            url = 'https://' + url
        async with session.get(url, timeout=timeout, headers=HEADERS, ssl=False) as r:
            if r.status == 200:
                return await r.text()
    except Exception:
        pass
    return ""


async def scrape_website_data(session, url, semaphore, crawl_contact=False):
    async with semaphore:
        html = await fetch_page(session, url)
        if not html and url.startswith('https://'):
            url2 = url.replace('https://', 'http://', 1)
            html = await fetch_page(session, url2)
            if html:
                url = url2
        all_html = html or ""
        if crawl_contact and html:
            p = urlparse(url)
            base = f"{p.scheme}://{p.netloc}"
            pages = await asyncio.gather(
                *[fetch_page(session, base + cp) for cp in CONTACT_PATHS],
                return_exceptions=True
            )
            for pg in pages:
                if isinstance(pg, str) and pg:
                    all_html += "\n" + pg

        emails = {e.lower() for e in re.findall(EMAIL_REGEX, all_html) if is_valid_email(e)}
        phones = set()
        for p in re.findall(PHONE_REGEX, all_html):
            c = normalize_phone(p)
            if c:
                phones.add(c)

        biz = get_clean_domain_name(url)
        return [{
            "Business Name": biz,
            "Phone Number": ", ".join(sorted(phones)) if phones else "N/A",
            "Website": url,
            "Official Email ID": ", ".join(sorted(emails)) if emails else "N/A",
            "Address": "N/A", "Rating": "N/A", "Source": "aiohttp"
        }]


# ---------------- API: DIAGNOSE ----------------
@app.get("/api/diagnose")
def diagnose():
    result = {"python": sys.version.split()[0], "selenium_launch": None, "page_title": None, "error": None}
    driver, err = launch_driver()
    if err:
        result["selenium_launch"] = "FAIL"
        result["error"] = err[:500]
        return result
    try:
        driver.get("https://www.google.com")
        time.sleep(2)
        result["selenium_launch"] = "OK"
        result["page_title"] = driver.title
    except Exception as e:
        result["selenium_launch"] = "FAIL"
        result["error"] = f"get() failed: {e}"
    finally:
        try:
            driver.quit()
        except Exception:
            pass
    return result


# ---------------- API: GMAPS ----------------
@app.post("/api/gmaps-scrape")
async def api_gmaps_scrape(gmaps_query: str = Form(...)):
    loop = asyncio.get_running_loop()
    try:
        raw = await asyncio.wait_for(
            loop.run_in_executor(None, run_selenium_gmaps, gmaps_query),
            timeout=240
        )
    except asyncio.TimeoutError:
        return {"status": "error", "total": 1,
                "data": error_row("Gmaps timeout (240s)", source="Timeout")}
    except Exception as e:
        return {"status": "error", "total": 1,
                "data": error_row(f"Gmaps exception: {e}", source="Exception")}

    with_site = [l for l in raw if is_valid_website(l.get("Website", ""))]
    without_site = [l for l in raw if l not in with_site]
    expanded = list(without_site)

    if with_site:
        sem = asyncio.Semaphore(15)
        conn = aiohttp.TCPConnector(ssl=False, limit=50)
        async with aiohttp.ClientSession(connector=conn) as s:
            tasks = [scrape_website_data(s, lead["Website"], sem) for lead in with_site]
            nested = await asyncio.gather(*tasks, return_exceptions=True)
            for res, lead in zip(nested, with_site):
                if isinstance(res, Exception) or not res:
                    expanded.append(lead)
                    continue
                for w in res:
                    merged_phone = lead["Phone Number"] if lead["Phone Number"] != "N/A" else w["Phone Number"]
                    merged_email = w["Official Email ID"] if w["Official Email ID"] != "N/A" else "N/A"
                    expanded.append({
                        "Business Name": lead["Business Name"],
                        "Phone Number": merged_phone,
                        "Website": lead["Website"],
                        "Official Email ID": merged_email,
                        "Address": lead["Address"],
                        "Rating": lead["Rating"],
                        "Source": "Gmaps + Deep Scan"
                    })

    if not expanded:
        expanded = error_row("No results", source="Gmaps")

    save_excel(expanded)
    log(f"GMAPS API: returned {len(expanded)} rows")
    return {"status": "success", "total": len(expanded), "data": expanded}


# ---------------- API: BULK ----------------
@app.post("/api/bulk-scrape")
async def api_bulk_scrape(urls_text: str = Form(...), crawl_contact: bool = Form(False)):
    urls = [u.strip() for u in urls_text.split('\n') if u.strip()]
    sem = asyncio.Semaphore(20)
    conn = aiohttp.TCPConnector(ssl=False, limit=50)
    async with aiohttp.ClientSession(connector=conn) as s:
        tasks = [scrape_website_data(s, u, sem, crawl_contact=crawl_contact) for u in urls]
        nested = await asyncio.gather(*tasks, return_exceptions=True)
    all_res = []
    for r in nested:
        if isinstance(r, list):
            all_res.extend(r)
    save_excel(all_res)
    return {"status": "success", "total": len(all_res), "data": all_res}


# ---------------- API: SINGLE WEB ----------------
@app.post("/api/web-scrape")
async def api_web_scrape(target_url: str = Form(...), use_selenium: bool = Form(False)):
    loop = asyncio.get_running_loop()
    results = []
    if not use_selenium:
        conn = aiohttp.TCPConnector(ssl=False)
        async with aiohttp.ClientSession(connector=conn) as s:
            sem = asyncio.Semaphore(1)
            results = await scrape_website_data(s, target_url, sem, crawl_contact=True)
        empty = all(r["Official Email ID"] == "N/A" and r["Phone Number"] == "N/A" for r in results)
        log(f"aiohttp: {len(results)} rows, empty={empty}")
        if empty:
            use_selenium = True

    if use_selenium:
        try:
            results = await asyncio.wait_for(
                loop.run_in_executor(None, run_selenium_web_scraper, target_url, True),
                timeout=180
            )
        except asyncio.TimeoutError:
            results = error_row("Selenium timeout (180s)", target_url, source="Timeout")
        except Exception as e:
            results = error_row(f"Selenium exception: {e}", target_url, source="Exception")

    save_excel(results)
    return {"status": "success", "total": len(results), "data": results}


# ---------------- MISC ----------------
@app.get("/download")
def download_file():
    if os.path.exists(EXCEL_FILE):
        return FileResponse(path=EXCEL_FILE, filename="Master_Leads_Report.xlsx",
                            media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    return {"error": "File not found"}


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def index():
    if os.path.exists("index.html"):
        with open("index.html", encoding="utf-8") as f:
            return f.read()
    return "<h1>API running</h1>"


if __name__ == "__main__":
    import uvicorn
    log(f"Starting — HEADLESS={HEADLESS}")
    uvicorn.run(app, host="0.0.0.0", port=3000)  
