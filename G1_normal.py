import re
import io
import sys
import time
import asyncio
import requests
import requests as req_lib
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from urllib.parse import quote
from functools import partial
import openpyxl
import calendar
import os

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, StaleElementReferenceException
)
from playwright.async_api import async_playwright


# ═══════════════════════════════════════════════════════════════════════
# LOG 攔截：把所有 print 同時寫到畫面 + 記憶體
# ═══════════════════════════════════════════════════════════════════════

class TeeLogger:
    """同時輸出到 stdout 和 StringIO buffer"""
    def __init__(self):
        self.buffer = io.StringIO()
        self._stdout = sys.stdout

    def write(self, msg):
        self._stdout.write(msg)
        self.buffer.write(msg)

    def flush(self):
        self._stdout.flush()

    def get_log(self):
        return self.buffer.getvalue()

logger = TeeLogger()
sys.stdout = logger


# ═══════════════════════════════════════════════════════════════════════
# STEP 1: 自動計算日期範圍
# ═══════════════════════════════════════════════════════════════════════

today = datetime.today()
day = today.day

if day == 1:
    # 上個月 21 號到月底
    first_of_this_month = today.replace(day=1)
    last_month = first_of_this_month - timedelta(days=1)
    last_day = calendar.monthrange(last_month.year, last_month.month)[1]
    start_dt = last_month.replace(day=21)
    end_dt   = last_month.replace(day=last_day)

elif day == 11:
    # 本月 1 號到 10 號
    start_dt = today.replace(day=1)
    end_dt   = today.replace(day=10)

elif day == 21:
    # 本月 11 號到 20 號
    start_dt = today.replace(day=11)
    end_dt   = today.replace(day=20)

else:
    # 非排程日，手動測試時可用，預設抓最近10天
    end_dt   = today
    start_dt = today - timedelta(days=10)
    print(f"  ⚠ 今天是 {day} 號，非標準排程日，使用預設區間（最近10天）")

# 把 start_dt / end_dt 的時間部分清零，避免時間比較問題
start_dt = start_dt.replace(hour=0, minute=0, second=0, microsecond=0)
end_dt   = end_dt.replace(hour=23, minute=59, second=59, microsecond=0)

print(f"  ✓ 時間區間: {start_dt.date()} ~ {end_dt.date()}")

# 所有機構共用的結果清單與已訪問 URL
all_results = []
visited_urls = set()


# ═══════════════════════════════════════════════════════════════════════
# 寄信功能
# ═══════════════════════════════════════════════════════════════════════

def send_email(excel_path: str, log_text: str):
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders

    sender    = os.environ.get("EMAIL_SENDER", "")
    password  = os.environ.get("EMAIL_PASSWORD", "")
    recipient = os.environ.get("EMAIL_RECIPIENT", "")

    if not sender or not password or not recipient:
        print("⚠ EMAIL_SENDER / EMAIL_PASSWORD / EMAIL_RECIPIENT 未設定，跳過寄信")
        return

    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = recipient
    msg["Subject"] = f"爬蟲結果 {today.strftime('%Y/%m/%d')}（區間 {start_dt.date()} ~ {end_dt.date()}）"

    body = (
        f"爬蟲執行完畢。\n"
        f"執行日期：{today.strftime('%Y/%m/%d %H:%M')}\n"
        f"資料區間：{start_dt.date()} ~ {end_dt.date()}\n"
        f"共收錄：{len(all_results)} 筆\n\n"
        f"詳細 log 請見附件 scraper_log.txt，Excel 結果請見附件。"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # 附件 1：Excel
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename={os.path.basename(excel_path)}")
            msg.attach(part)

    # 附件 2：log 文字檔
    log_bytes = log_text.encode("utf-8")
    log_part = MIMEBase("text", "plain")
    log_part.set_payload(log_bytes)
    encoders.encode_base64(log_part)
    log_part.add_header("Content-Disposition", "attachment; filename=scraper_log.txt")
    msg.attach(log_part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, password)
            server.sendmail(sender, recipient, msg.as_string())
        print(f"✅ Email 已寄出至 {recipient}")
    except Exception as e:
        print(f"❌ 寄信失敗：{e}")


# ═══════════════════════════════════════════════════════════════════════
# 共用：建立 headless Chrome Options
# ═══════════════════════════════════════════════════════════════════════

def build_chrome_options():
    opts = Options()
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--window-size=1920,1080")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    return opts


# ═══════════════════════════════════════════════════════════════════════
# AC Global China Hub Scraper Class
# ═══════════════════════════════════════════════════════════════════════

class ACGlobalChinaHubScraper:

    TARGET_URL      = "https://www.atlanticcouncil.org/programs/global-china-hub/global-china-hub-publications/"
    KEYWORDS        = ["China", "Taiwan", "Taipei"]
    CARD_SEL        = "div.gta-post-embed--container"
    SECTION_KEYWORDS = ["reports", "issue brief", "new atlanticist"]

    def __init__(self, start_dt, end_dt, visited_urls):
        self.start_dt     = start_dt
        self.end_dt       = end_dt
        self.visited_urls = visited_urls
        self.results      = []
        self.driver       = None

    def setup_browser(self):
        self.driver = webdriver.Chrome(options=build_chrome_options())

    def parse_article_date(self, date_str):
        if not date_str:
            return None
        date_str = date_str.strip()
        for fmt in ("%B %d, %Y", "%B %Y", "%b %d, %Y", "%b %Y",
                    "%b. %d, %Y", "%d %B %Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                pass
        return None

    def dismiss_cookie_banner(self):
        try:
            for btn in self.driver.find_elements(By.TAG_NAME, "button"):
                try:
                    txt = btn.text.strip().lower()
                    if any(w in txt for w in ["accept", "agree", "got it"]):
                        self.driver.execute_script("arguments[0].click();", btn)
                        time.sleep(1)
                        return
                except Exception:
                    pass
        except Exception:
            pass

    def scroll_to_element(self, el):
        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", el)
        time.sleep(2)

    def wait_for_cards(self, timeout=15):
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, self.CARD_SEL))
            )
        except TimeoutException:
            print("    ⚠ 等待卡片逾時，嘗試繼續...")
        time.sleep(1)

    def find_section_h2s(self):
        all_h2 = self.driver.find_elements(By.TAG_NAME, "h2")
        result = []
        for h2 in all_h2:
            try:
                txt = h2.text.strip().lower()
                if not txt:
                    continue
                in_noise = self.driver.execute_script("""
                    var el = arguments[0];
                    return !!(el.closest('nav') || el.closest('header') || el.closest('footer'));
                """, h2)
                if in_noise:
                    continue
                for kw in self.SECTION_KEYWORDS:
                    if kw in txt:
                        result.append((h2.text.strip(), h2))
                        break
            except Exception:
                continue
        return result

    def get_cards_in_section(self, h2_el, next_h2_el=None):
        self.scroll_to_element(h2_el)

        h2_top = self.driver.execute_script("""
            var el = arguments[0], top = 0;
            while (el) { top += el.offsetTop; el = el.offsetParent; }
            return top;
        """, h2_el)

        next_top = self.driver.execute_script("""
            var el = arguments[0];
            if (!el) return 999999999;
            var top = 0;
            while (el) { top += el.offsetTop; el = el.offsetParent; }
            return top;
        """, next_h2_el) if next_h2_el else 999999999

        try:
            WebDriverWait(self.driver, 15).until(lambda d: self.driver.execute_script("""
                var h2Top   = arguments[0];
                var nextTop = arguments[1];
                var cards   = document.querySelectorAll('div.gta-post-embed--container');
                for (var i = 0; i < cards.length; i++) {
                    var el = cards[i], top = 0;
                    while (el) { top += el.offsetTop; el = el.offsetParent; }
                    if (top >= h2Top && top < nextTop) return true;
                }
                return false;
            """, h2_top, next_top))
        except TimeoutException:
            print(f"    ⚠ 等待 [{h2_el.text.strip()}] 卡片逾時，嘗試繼續...")
        time.sleep(1)

        return self.driver.execute_script("""
            var h2Top   = arguments[0];
            var nextTop = arguments[1];

            function getTop(el) {
                var top = 0;
                while (el) { top += el.offsetTop; el = el.offsetParent; }
                return top;
            }

            var cards = [];
            var seen  = new Set();
            document.querySelectorAll('div.gta-post-embed--container').forEach(function(card) {
                var top = getTop(card);
                if (top < h2Top || top >= nextTop) return;

                var linkEl = card.querySelector('a[href*="atlanticcouncil"]');
                if (!linkEl) return;
                var url = linkEl.href;
                if (seen.has(url)) return;
                seen.add(url);

                var titleEl = card.querySelector('h4, h3, h2');
                var title   = titleEl ? titleEl.textContent.trim() : '';

                var dateStr = '';
                var timeEl  = card.querySelector('time[datetime]');
                if (timeEl) {
                    dateStr = timeEl.getAttribute('datetime') || timeEl.textContent.trim();
                } else {
                    var dateEl = card.querySelector(
                        'span.gta-post-embed--heading--date, span[class*="date"], [class*="date"]'
                    );
                    if (dateEl) dateStr = dateEl.textContent.trim();
                }
                cards.push({title: title, url: url, date_str: dateStr});
            });

            var loadMoreUrl = null;
            document.querySelectorAll('a').forEach(function(a) {
                if (loadMoreUrl) return;
                var top = getTop(a);
                if (top >= h2Top && top < nextTop &&
                    a.textContent.trim().toLowerCase() === 'load more') {
                    loadMoreUrl = a.href;
                }
            });

            return {cards: cards, load_more_url: loadMoreUrl};
        """, h2_top, next_top)

    def fetch_page_cards(self, url):
        self.driver.execute_script("window.open('');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        cards   = []
        next_lm = None
        try:
            self.driver.get(url)
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, self.CARD_SEL))
                )
            except TimeoutException:
                print("    ⚠ 等待卡片逾時，嘗試繼續...")
            time.sleep(1)

            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            self.driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(1)

            cards = self.driver.execute_script("""
                var result = [];
                var seen   = new Set();
                document.querySelectorAll('div.gta-post-embed--container').forEach(function(card) {
                    var linkEl = card.querySelector('a[href*="atlanticcouncil"]');
                    if (!linkEl) return;
                    var url = linkEl.href;
                    if (seen.has(url)) return;
                    seen.add(url);

                    var titleEl = card.querySelector('h4, h3, h2');
                    var title   = titleEl ? titleEl.textContent.trim() : '';

                    var dateStr = '';
                    var timeEl  = card.querySelector('time[datetime]');
                    if (timeEl) {
                        dateStr = timeEl.getAttribute('datetime') || timeEl.textContent.trim();
                    } else {
                        var dateEl = card.querySelector(
                            'span.gta-post-embed--heading--date, span[class*="date"], [class*="date"]'
                        );
                        if (dateEl) dateStr = dateEl.textContent.trim();
                    }
                    result.push({title: title, url: url, date_str: dateStr});
                });
                return result;
            """)

            try:
                for a in self.driver.find_elements(By.TAG_NAME, "a"):
                    if a.text.strip().lower() == "load more":
                        next_lm = a.get_attribute("href")
                        break
            except Exception:
                pass

            return cards, next_lm

        except Exception as e:
            print(f"    ✗ fetch_page_cards 失敗: {e}")
            return [], None
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    def check_article(self, url):
        self.driver.execute_script("window.open('');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        try:
            self.driver.get(url)
            try:
                WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "h1"))
                )
            except TimeoutException:
                pass
            time.sleep(1)

            title = ""
            for sel in ["h1.gta-site-banner--title", "h1"]:
                try:
                    t = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                    if t:
                        title = t
                        break
                except NoSuchElementException:
                    pass

            body_text = title
            try:
                body_text += " " + self.driver.execute_script("""
                    var noiseSelectors = [
                        '[class*="gta-expert-embed"]',
                        '[class*="gta-horizontal-featured"]',
                        '[class*="gta-section-opener"]',
                        '[class*="related"]', '[class*="explore"]',
                        '[class*="recommended"]',
                        'aside', 'footer', 'nav', 'header'
                    ];
                    var root = document.body.cloneNode(true);
                    noiseSelectors.forEach(function(sel){
                        root.querySelectorAll(sel).forEach(function(el){ el.remove(); });
                    });
                    return root.innerText || root.textContent || '';
                """)
            except Exception:
                try:
                    body_text += " " + self.driver.find_element(By.TAG_NAME, "body").text
                except Exception:
                    pass

            found_keywords = [kw for kw in self.KEYWORDS if re.search(kw, body_text, re.IGNORECASE)]

            has_pdf = False
            try:
                for btn in self.driver.find_elements(
                    By.CSS_SELECTOR, "a.wp-block-button__link, a.wp-element-button"
                ):
                    if "pdf" in btn.text.lower():
                        has_pdf = True
                        break
            except Exception:
                pass

            return found_keywords, has_pdf, title

        except Exception as e:
            print(f"      ✗ 無法讀取文章: {e}")
            return [], False, ""
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    def process_section_data(self, label, initial_cards, first_load_more_url):
        print(f"\n{'-'*20}")
        print(f"類別: {label}")
        print(f"{'-'*20}")

        consecutive_early     = 0
        global_idx            = 0
        current_load_more_url = first_load_more_url
        all_batches           = [initial_cards]
        batch_no              = 0

        while batch_no < len(all_batches):
            batch    = all_batches[batch_no]
            batch_no += 1

            if not batch:
                if current_load_more_url:
                    print(f"  → 載入更多: {current_load_more_url}")
                    extra_cards, next_url = self.fetch_page_cards(current_load_more_url)
                    current_load_more_url = next_url
                    if extra_cards:
                        all_batches.append(extra_cards)
                continue

            stop_flag = False
            for card in batch:
                global_idx += 1
                art_url   = card.get("url", "")
                art_title = card.get("title", "")
                date_str_ = card.get("date_str", "")
                art_date  = self.parse_article_date(date_str_)

                if not art_date or not art_url:
                    print(f"    [{global_idx}] (無日期/URL) {art_title[:50]} — 跳過")
                    continue

                date_display = art_date.strftime("%Y/%m/%d")

                if art_date > self.end_dt:
                    print(f"    [{global_idx}] {date_display} | {art_title[:55]} - 晚於結束時間，跳過")
                    consecutive_early = 0
                    continue

                if art_date < self.start_dt:
                    consecutive_early += 1
                    print(f"    [{global_idx}] {date_display} | {art_title[:55]} - 早於起始時間 (連續: {consecutive_early})")
                    if consecutive_early >= 2:
                        print(f"  → 連續2篇早於起始時間，停止 [{label}]")
                        stop_flag = True
                        break
                    continue

                consecutive_early = 0

                if art_url in self.visited_urls:
                    print(f"    [{global_idx}] {date_display} | {art_title[:55]} - 已拜訪，跳過")
                    continue

                self.visited_urls.add(art_url)
                print(f"    [{global_idx}] {date_display} | {art_title[:55]}")
                print(f"         → 進入文章: {art_url}")

                found_kws, has_pdf, full_title = self.check_article(art_url)

                if found_kws and has_pdf:
                    keywords_str = ", ".join(kw for kw in self.KEYWORDS if kw in found_kws)
                    self.results.append({
                        "機構":   "Atlantic Council - Global China Hub",
                        "日期":   date_display,
                        "標題":   full_title or art_title,
                        "網址":   art_url,
                        "關鍵字": keywords_str,
                    })
                    print(f"      ✓ 收錄! 關鍵字: {keywords_str}")
                else:
                    reason = []
                    if not found_kws: reason.append("無關鍵字")
                    if not has_pdf:   reason.append("無PDF")
                    print(f"      ✗ 不符合 ({', '.join(reason)})")

            if stop_flag:
                return

            if batch_no >= len(all_batches):
                if current_load_more_url:
                    print(f"  → 載入更多: {current_load_more_url}")
                    extra_cards, next_url = self.fetch_page_cards(current_load_more_url)
                    current_load_more_url = next_url
                    if extra_cards:
                        all_batches.append(extra_cards)
                    else:
                        print(f"  → 沒有更多文章，[{label}] 完成")
                else:
                    print(f"  → 沒有更多文章，[{label}] 完成")

    def run(self):
        print("=" * 60)
        print("Atlantic Council – Global China Hub 文章爬蟲")
        print("=" * 60)

        self.setup_browser()
        self.driver.get(self.TARGET_URL)
        time.sleep(3)
        self.dismiss_cookie_banner()
        time.sleep(1)

        print("\n正在捲動頁面讓所有 section 載入...")
        scroll_height = self.driver.execute_script("return document.body.scrollHeight")
        step = scroll_height // 6
        for i in range(1, 7):
            self.driver.execute_script(f"window.scrollTo(0, {step * i});")
            time.sleep(1.5)
        self.driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(2)

        print("正在解析頁面 section 結構...")
        section_h2s = self.find_section_h2s()

        if not section_h2s:
            print("  ✗ 未找到任何 section，請檢查頁面結構")
            self.driver.quit()
            return self.results

        print(f"  找到 {len(section_h2s)} 個 section: {[l for l, _ in section_h2s]}")

        sections_info = []
        for i, (label, h2_el) in enumerate(section_h2s):
            next_h2_el = section_h2s[i + 1][1] if i + 1 < len(section_h2s) else None
            print(f"  正在捲動到 [{label}]...")
            sec_data = self.get_cards_in_section(h2_el, next_h2_el)
            sections_info.append({
                "label":         label,
                "card_data":     sec_data["cards"],
                "load_more_url": sec_data["load_more_url"],
            })
            print(f"  ✓ 找到類別: {label} "
                  f"({len(sec_data['cards'])} 篇, Load More: {sec_data['load_more_url']})")

        self.driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(1)

        for sec in sections_info:
            self.process_section_data(
                label               = sec["label"],
                initial_cards       = sec["card_data"],
                first_load_more_url = sec["load_more_url"],
            )

        print(f"AC Global China Hub 爬取完成！共收錄 {len(self.results)} 篇文章")
        self.driver.quit()
        print("✓ GCH 瀏覽器已關閉")
        return self.results


# ═══════════════════════════════════════════════════════════════════════
# Atlantic Council Scraper Class
# ═══════════════════════════════════════════════════════════════════════

class AtlanticCouncilScraper:

    SEARCH_QUERIES = ["china", "taiwan", "taipei"]
    FACETS = [
        "focus%3A+Content+Series",
        "focus%3A+Issue+briefs+and+reports"
    ]
    KEYWORDS = ["China", "Taiwan", "Taipei"]

    def __init__(self, start_dt, end_dt, visited_urls):
        self.start_dt = start_dt
        self.end_dt = end_dt
        self.visited_urls = visited_urls
        self.results = []
        self.driver = None
        self.wait = None

    def setup_browser(self):
        self.driver = webdriver.Chrome(options=build_chrome_options())
        self.wait = WebDriverWait(self.driver, 20)

    def parse_article_date(self, date_str):
        date_str = date_str.strip()
        for fmt in ("%B %d, %Y", "%B %Y", "%b %d, %Y", "%b %Y"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                pass
        return None

    def dismiss_cookie_banner(self):
        try:
            for btn in self.driver.find_elements(By.TAG_NAME, "button"):
                try:
                    txt = btn.text.strip().lower()
                    if any(w in txt for w in ["accept", "agree", "got it"]):
                        self.driver.execute_script("arguments[0].click();", btn)
                        time.sleep(1)
                        return
                except Exception:
                    pass
        except Exception:
            pass

    def select_past_year(self):
        self.dismiss_cookie_banner()
        try:
            selects = self.driver.find_elements(By.TAG_NAME, "select")
            for sel_el in selects:
                opts = sel_el.find_elements(By.TAG_NAME, "option")
                for opt in opts:
                    if "past year" in opt.text.strip().lower():
                        self.driver.execute_script(
                            "arguments[0].selected = true; "
                            "arguments[0].parentElement.dispatchEvent(new Event('change', {bubbles:true}));",
                            opt
                        )
                        print("    ✓ 已選擇 Past year，等待 5 秒讓文章載入...")
                        time.sleep(5)
                        print(f"    ✓ 篩選後 URL: {self.driver.current_url}")
                        return
            print("    ✗ 找不到 Past year 選項")
        except Exception as e:
            print(f"    ✗ select_past_year 失敗: {e}")

    def get_all_cards(self):
        try:
            return self.driver.find_elements(By.CSS_SELECTOR, "article.record")
        except Exception:
            return []

    def click_load_more(self):
        selectors = [
            "a.j-posts--loadmore",
            "a.o-archives--pagination-button",
            "button.btn-load-more",
            "a[class*='loadmore']",
            "button[class*='load-more']",
            "a.s-button.j-posts--loadmore",
        ]
        for sel in selectors:
            try:
                btn = self.driver.find_element(By.CSS_SELECTOR, sel)
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                time.sleep(0.5)
                self.driver.execute_script("arguments[0].click();", btn)
                time.sleep(4)
                return True
            except NoSuchElementException:
                continue
            except Exception as e:
                print(f"    ✗ click_load_more 失敗: {e}")
                return False
        return False

    def get_card_date(self, card):
        try:
            date_el = card.find_element(By.CSS_SELECTOR, "h5.record__meta--date")
            return self.parse_article_date(date_el.text)
        except NoSuchElementException:
            return None

    def get_card_url(self, card):
        try:
            link = card.find_element(By.CSS_SELECTOR, "a")
            return link.get_attribute("href")
        except NoSuchElementException:
            return None

    def get_card_title(self, card):
        try:
            for sel in ["h4 a", "h3 a", "a.seon-lookup-checked", "a"]:
                try:
                    el = card.find_element(By.CSS_SELECTOR, sel)
                    t = el.text.strip()
                    if t:
                        return t
                except NoSuchElementException:
                    continue
        except Exception:
            pass
        return ""

    def get_report_card_date(self, card):
        for sel in [
            "span.gta-post-embed--heading--date",
            "span[class*='date']",
            "time",
            "p[class*='date']",
        ]:
            try:
                el = card.find_element(By.CSS_SELECTOR, sel)
                txt = el.get_attribute("datetime") or el.text.strip()
                dt = self.parse_article_date(txt)
                if dt:
                    return dt
            except NoSuchElementException:
                continue
        return None

    def get_report_card_title(self, card):
        try:
            for sel in ["h4.gta-post-embed--title", "h4", "h3"]:
                try:
                    el = card.find_element(By.CSS_SELECTOR, sel)
                    t = el.text.strip()
                    if t:
                        return t
                except NoSuchElementException:
                    continue
        except Exception:
            pass
        return ""

    def get_report_card_url(self, card):
        for sel in ["h4 a", "h3 a", "a[href*='/in-depth']", "a[href*='/content-series']", "a"]:
            try:
                el = card.find_element(By.CSS_SELECTOR, sel)
                href = el.get_attribute("href")
                if href and "atlanticcouncil.org" in href:
                    return href
            except NoSuchElementException:
                continue
        return None

    def check_article(self, url):
        self.driver.execute_script("window.open('');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        try:
            self.driver.get(url)
            time.sleep(2)

            title = ""
            for sel in ["h1.gta-site-banner--title", "h1"]:
                try:
                    t = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                    if t:
                        title = t
                        break
                except NoSuchElementException:
                    pass

            body_text = title
            try:
                body_text += " " + self.driver.execute_script("""
                    var noiseSelectors = [
                        '[class*="gta-expert-embed"]',
                        '[class*="gta-horizontal-featured"]',
                        '[class*="gta-section-opener"]',
                        '[class*="related"]',
                        '[class*="explore"]',
                        '[class*="recommended"]',
                        'aside', 'footer', 'nav', 'header'
                    ];
                    var root = document.body.cloneNode(true);
                    noiseSelectors.forEach(function(sel){
                        root.querySelectorAll(sel).forEach(function(el){ el.remove(); });
                    });
                    return root.innerText || root.textContent || '';
                """)
            except Exception:
                try:
                    body_text += " " + self.driver.find_element(By.TAG_NAME, "body").text
                except Exception:
                    pass

            found_keywords = [kw for kw in self.KEYWORDS if re.search(kw, body_text, re.IGNORECASE)]

            has_pdf = False
            try:
                buttons = self.driver.find_elements(
                    By.CSS_SELECTOR, "a.wp-block-button__link, a.wp-element-button"
                )
                for btn in buttons:
                    if "pdf" in btn.text.lower():
                        has_pdf = True
                        break
            except Exception:
                pass

            return found_keywords, has_pdf, title

        except Exception as e:
            print(f"      ✗ 無法讀取文章: {e}")
            return [], False, ""
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    def run(self):
        print("=" * 60)
        print("Atlantic Council 文章爬蟲")
        print("=" * 60)

        self.setup_browser()

        for query in self.SEARCH_QUERIES:
            for facet in self.FACETS:
                facet_label = "Content Series" if "Content+Series" in facet else "Issue briefs and reports"
                print(f"\n{'-'*20}")
                print(f"查詢: {query}  |  分類: {facet_label}")
                print(f"{'-'*20}")

                first_url = (
                    f"https://www.atlanticcouncil.org/search/?query={query}"
                    f"&page=0&facetFilters={facet}&numericFilters="
                )
                print(f"\n  → 載入搜尋頁: {first_url}")
                self.driver.get(first_url)
                time.sleep(3)
                self.select_past_year()

                processed_count = 0
                consecutive_early = 0
                stop_this = False

                while True:
                    all_cards = self.get_all_cards()
                    new_cards = all_cards[processed_count:]

                    if not new_cards:
                        has_more = self.click_load_more()
                        if not has_more:
                            print("  → 沒有更多文章可載入")
                            break
                        continue

                    print(f"\n  → 處理第 {processed_count+1}~{processed_count+len(new_cards)} 筆文章")

                    for i, card in enumerate(new_cards):
                        global_idx = processed_count + i + 1
                        try:
                            art_date = self.get_card_date(card)
                            art_url  = self.get_card_url(card)
                        except StaleElementReferenceException:
                            all_cards = self.get_all_cards()
                            idx = processed_count + i
                            if idx >= len(all_cards):
                                break
                            art_date = self.get_card_date(all_cards[idx])
                            art_url  = self.get_card_url(all_cards[idx])

                        if not art_date or not art_url:
                            continue

                        date_str_display = art_date.strftime("%Y/%m/%d")

                        if art_date > self.end_dt:
                            card_title = self.get_card_title(card)
                            print(f"    [{global_idx}] {date_str_display} | {card_title[:60]} - 晚於結束時間，跳過")
                            consecutive_early = 0
                            continue

                        if art_date < self.start_dt:
                            consecutive_early += 1
                            card_title = self.get_card_title(card)
                            print(f"    [{global_idx}] {date_str_display} | {card_title[:60]} - 早於起始時間 (連續: {consecutive_early})")
                            if consecutive_early >= 5:
                                print("  → 連續5篇早於起始時間，停止此查詢/分類")
                                stop_this = True
                                break
                            continue

                        consecutive_early = 0

                        if art_url in self.visited_urls:
                            card_title = self.get_card_title(card)
                            print(f"    [{global_idx}] {date_str_display} | {card_title[:60]} - 已拜訪，跳過")
                            continue

                        self.visited_urls.add(art_url)
                        card_title = self.get_card_title(card)
                        print(f"    [{global_idx}] {date_str_display} | {card_title[:60]}")
                        print(f"         → 進入文章: {art_url}")

                        found_kws, has_pdf, title = self.check_article(art_url)

                        if found_kws and has_pdf:
                            keywords_str = ", ".join(kw for kw in self.KEYWORDS if kw in found_kws)
                            self.results.append({
                                "機構": "Atlantic Council",
                                "日期":  art_date.strftime("%Y/%m/%d"),
                                "標題":  title,
                                "網址":  art_url,
                                "關鍵字": keywords_str,
                            })
                            print(f"      ✓ 收錄! 關鍵字: {keywords_str}")
                        else:
                            reason = []
                            if not found_kws: reason.append("無關鍵字")
                            if not has_pdf:   reason.append("無PDF")
                            print(f"      ✗ 不符合 ({', '.join(reason)})")

                    processed_count = len(self.get_all_cards())

                    if stop_this:
                        break

                    has_more = self.click_load_more()
                    if not has_more:
                        print("  → 沒有更多文章，此查詢/分類完成")
                        break

        # ── 額外來源：in-depth-research-reports ──────────────────────
        sep = "-" * 20
        print(f"\n{sep}")
        print("額外來源: in-depth-research-reports")
        print(sep)

        REPORTS_URL = "https://www.atlanticcouncil.org/in-depth-research-reports/"
        print(f"  → 載入: {REPORTS_URL}")
        self.driver.get(REPORTS_URL)
        time.sleep(3)

        time.sleep(3)
        REPORT_CARD_SEL = "div.gta-post-embed--container"

        processed_count_r = 0
        consecutive_early_r = 0
        stop_reports = False

        while True:
            all_cards_r = self.driver.find_elements(By.CSS_SELECTOR, REPORT_CARD_SEL)
            new_cards_r = all_cards_r[processed_count_r:]

            if not new_cards_r:
                has_more = self.click_load_more()
                if not has_more:
                    print("  → 沒有更多文章可載入")
                    break
                time.sleep(3)
                continue

            print(f"  → 處理第 {processed_count_r+1}~{processed_count_r+len(new_cards_r)} 筆文章")

            for i, card in enumerate(new_cards_r):
                global_idx = processed_count_r + i + 1
                try:
                    art_date     = self.get_report_card_date(card)
                    art_url      = self.get_report_card_url(card)
                    card_title_r = self.get_report_card_title(card)
                except StaleElementReferenceException:
                    all_cards_r = self.driver.find_elements(By.CSS_SELECTOR, REPORT_CARD_SEL)
                    idx = processed_count_r + i
                    if idx >= len(all_cards_r):
                        break
                    art_date     = self.get_report_card_date(all_cards_r[idx])
                    art_url      = self.get_report_card_url(all_cards_r[idx])
                    card_title_r = self.get_report_card_title(all_cards_r[idx])

                if not art_date or not art_url:
                    continue

                date_str_display = art_date.strftime("%Y/%m/%d")

                if art_date > self.end_dt:
                    print(f"    [{global_idx}] {date_str_display} | {card_title_r[:60]} - 晚於結束時間，跳過")
                    consecutive_early_r = 0
                    continue

                if art_date < self.start_dt:
                    consecutive_early_r += 1
                    print(f"    [{global_idx}] {date_str_display} | {card_title_r[:60]} - 早於起始時間 (連續: {consecutive_early_r})")
                    if consecutive_early_r >= 5:
                        print("  → 連續5篇早於起始時間，停止 in-depth-research-reports")
                        stop_reports = True
                        break
                    continue

                consecutive_early_r = 0

                if art_url in self.visited_urls:
                    print(f"    [{global_idx}] {date_str_display} | {card_title_r[:60]} - 已拜訪，跳過")
                    continue

                self.visited_urls.add(art_url)
                print(f"    [{global_idx}] {date_str_display} | {card_title_r[:60]}")
                print(f"         → 進入文章: {art_url}")

                found_kws, has_pdf, title = self.check_article(art_url)

                if found_kws and has_pdf:
                    keywords_str = ", ".join(kw for kw in self.KEYWORDS if kw in found_kws)
                    self.results.append({
                        "機構": "Atlantic Council",
                        "日期":  art_date.strftime("%Y/%m/%d"),
                        "標題":  title,
                        "網址":  art_url,
                        "關鍵字": keywords_str,
                    })
                    print(f"      ✓ 收錄! 關鍵字: {keywords_str}")
                else:
                    reason = []
                    if not found_kws: reason.append("無關鍵字")
                    if not has_pdf:   reason.append("無PDF")
                    print(f"      ✗ 不符合 ({', '.join(reason)})")

            processed_count_r = len(self.driver.find_elements(By.CSS_SELECTOR, REPORT_CARD_SEL))

            if stop_reports:
                break

            has_more = self.click_load_more()
            if not has_more:
                print("  → 沒有更多文章，in-depth-research-reports 完成")
                break

        print(f"Atlantic Council 爬取完成！共收錄 {len(self.results)} 篇文章")
        self.driver.quit()
        print("✓ AC 瀏覽器已關閉")
        return self.results


# ═══════════════════════════════════════════════════════════════════════
# RUSI Scraper Class
# ═══════════════════════════════════════════════════════════════════════

class RUSIScraper:

    INSTITUTION    = "The Royal United Services Institute"
    ALGOLIA_APP_ID = "5W5LULX8XD"
    ALGOLIA_API_KEY= "9d39b50e2f00b7d3f2bf638a47cc924b"
    ALGOLIA_HOST   = "https://5w5lulx8xd-dsn.algolia.net"
    INDEX_NAME     = "global_by_date"
    BASE_URL       = "https://www.rusi.org"
    KEYWORDS       = ["china", "taiwan", "taipei"]
    KEYWORD_DISPLAY= {"china": "China", "taiwan": "Taiwan", "taipei": "Taipei"}
    SKIP_TYPE_KEYWORDS = ["podcast", "recording"]
    HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept-Language": "en-US,en;q=0.9",
    }

    def __init__(self, start_dt, end_dt, visited_urls):
        self.start_dt = start_dt
        self.end_dt = end_dt
        self.visited_urls = visited_urls
        self.results = []

    def algolia_search(self, keyword: str, page: int = 0, hits_per_page: int = 100) -> dict:
        url = f"{self.ALGOLIA_HOST}/1/indexes/*/queries"
        params = {
            "x-algolia-agent": "Algolia for JavaScript (4.25.3); Browser; JS Helper (3.14.0)"
        }
        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "X-Algolia-Api-Key": self.ALGOLIA_API_KEY,
            "X-Algolia-Application-Id": self.ALGOLIA_APP_ID,
            "Origin": "https://www.rusi.org",
            "Referer": "https://www.rusi.org/",
            "User-Agent": "Mozilla/5.0",
        }
        payload = {
            "requests": [{
                "indexName": self.INDEX_NAME,
                "params": "&".join([
                    f"query={quote(keyword)}",
                    f"hitsPerPage={hits_per_page}",
                    f"page={page}",
                    "facets=%5B%22content_type%22%2C%22region_name%22%2C%22topics_name%22%5D",
                    "filters=" + quote(
                        "NOT type:'external_signpost' AND NOT efs:true AND ("
                        'content_type:"Briefing Papers" OR content_type:"Commentary" OR '
                        'content_type:"Conference Reports" OR content_type:"Emerging Insights" OR '
                        'content_type:"Occasional Papers" OR content_type:"Policy Briefs" OR '
                        'content_type:"RUSI Defence Systems" OR content_type:"RUSI Journal" OR '
                        'content_type:"RUSI Newsbrief" OR content_type:"Special Resources" OR '
                        'content_type:"Whitehall Papers" OR content_type:"Whitehall Reports" OR '
                        'content_type:"RUSI News" OR content_type:"Video Commentary" OR '
                        'content_type:"A Call to Arms Videos" OR content_type:"Decoding Counterterrorism Podcasts" OR '
                        'content_type:"External Publications" OR content_type:"Greening Defence Podcasts" OR '
                        'content_type:"Mind the Gulf Podcasts" OR content_type:"Research Event Recordings" OR '
                        'content_type:"RUSI Journal Radio" OR content_type:"RUSI Reflects" OR '
                        'content_type:"Strategic Hub for Organised Crime Research (SHOC)" OR '
                        'content_type:"Transatlantic Dialogue on China" OR content_type:"RUSI Annual Reports" OR '
                        'content_type:"Suspicious Transaction Report Podcasts" OR content_type:"War in Space Podcasts" OR '
                        'content_type:"Western Way of War Podcasts" OR content_type:"RUSI Analysis Podcasts" OR '
                        'content_type:"Adversarial Studies Videos" OR content_type:"Financial Crime Insights Podcasts" OR '
                        'content_type:"Members Event Recordings" OR content_type:"Global Security Briefing Podcasts" OR '
                        'content_type:"In Context Podcasts" OR content_type:"On the Cusp Podcasts" OR '
                        'content_type:"Talking Strategy Podcasts" OR content_type:"Environmental Security Podcasts" OR '
                        'content_type:"Bridging the Oceans Podcasts" OR content_type:"Projects" OR '
                        'content_type:"Explainers" OR content_type:"Disorder Podcast" OR '
                        'content_type:"Research Papers" OR content_type:"Insights Papers" OR '
                        'profile_type_name:"Staff" OR profile_type_name:"Associate Fellows" OR '
                        'profile_type_name:"Senior Associate Fellows" OR profile_type_name:"Distinguished Fellows" OR '
                        'profile_type_name:"Consultants" OR profile_type_name:"Trustees" OR '
                        'profile_type_name:"Advisory Board" OR profile_type_name:"Visiting Fellows" OR '
                        'profile_type_name:"Affiliate Experts")'
                    ),
                    "tagFilters=",
                ])
            }]
        }
        resp = requests.post(url, params=params, headers=headers, json=payload, timeout=20)
        resp.raise_for_status()
        return resp.json()

    def get_articles_in_range(self, keyword: str) -> list:
        collected = []
        consecutive_early = 0
        page = 0

        while True:
            try:
                data = self.algolia_search(keyword, page=page, hits_per_page=100)
            except Exception as ex:
                print(f"  ⚠️  Algolia 錯誤（第{page+1}頁）：{ex}")
                break

            results = data.get("results", [{}])[0]
            hits     = results.get("hits", [])
            nb_pages = results.get("nbPages", 1)

            if not hits:
                break

            stop = False
            for hit in hits:
                ctype = hit.get("content_type", "")
                if any(sk in ctype.lower() for sk in self.SKIP_TYPE_KEYWORDS):
                    continue

                if hit.get("premium", False):
                    continue

                ts = hit.get("authored_date")
                article_date = datetime.utcfromtimestamp(ts) if ts else None

                slug = hit.get("url", "")
                if not slug:
                    continue
                article_url = self.BASE_URL + slug if slug.startswith("/") else slug
                title = hit.get("title", "")

                if article_date is None:
                    pass
                elif article_date > self.end_dt:
                    continue
                elif article_date < self.start_dt:
                    consecutive_early += 1
                    print(f"    ⏩ {article_date.strftime('%Y/%m/%d')} 早於範圍（連續 {consecutive_early}/3）")
                    if consecutive_early >= 3:
                        stop = True
                        break
                    continue
                else:
                    consecutive_early = 0

                collected.append({
                    "url": article_url,
                    "date": article_date,
                    "title": title,
                    "content_type": ctype,
                })

            print(f"    第 {page+1} 頁：目前累計 {len(collected)} 篇")

            if stop or page >= nb_pages - 1:
                break
            page += 1
            time.sleep(0.3)

        return collected

    def check_article(self, url: str, target_keywords: list) -> dict:
        try:
            resp = requests.get(url, headers=self.HEADERS, timeout=20)
            resp.raise_for_status()
        except Exception as ex:
            print(f"  ⚠️  無法取得 {url}：{ex}")
            return None

        soup = BeautifulSoup(resp.text, "html.parser")

        page_text_upper = soup.get_text(" ", strip=True).upper()
        if "MEMBERS ONLY" in page_text_upper:
            members_badge = soup.find(
                lambda tag: tag.name in ["span", "div", "p", "li"] and
                "MEMBERS ONLY" in tag.get_text(strip=True).upper() and
                len(tag.get_text(strip=True)) < 50
            )
            if members_badge:
                return {"skip_reason": "members_only"}

        title_block = soup.find(class_=lambda c: c and "TitleBlock-module" in c)
        h1 = title_block.find("h1") if title_block else soup.find("h1")
        title = h1.get_text(strip=True) if h1 else ""

        body_parts = []
        for cls_kw in ["LeadParagraph-module", "Article-module--mainBody", "Article-module--paragraphs", "Article-module--contentArea"]:
            el = soup.find(class_=lambda c: c and cls_kw in c)
            if el:
                body_parts.append(el.get_text(" ", strip=True))
        if body_parts:
            full_text = " ".join(body_parts)
        else:
            for tag in soup(["nav", "footer", "script", "style", "header"]):
                tag.decompose()
            full_text = soup.get_text(" ", strip=True)

        found = []
        for kw in target_keywords:
            if re.search(re.escape(kw), full_text, re.IGNORECASE):
                found.append(self.KEYWORD_DISPLAY[kw.lower()])

        has_pdf = False
        for a in soup.find_all("a", href=True):
            href       = a.get("href", "")
            cls        = " ".join(a.get("class", []))
            aria_label = a.get("aria-label", "").lower()

            is_download_attr  = a.has_attr("download") and href.lower().endswith(".pdf")
            is_download_class = "DownloadFile-module" in cls or "downloadFile" in cls.lower()
            is_aria_download  = "download" in aria_label and href.lower().endswith(".pdf")

            if is_download_attr or is_download_class or is_aria_download:
                has_pdf = True
                break

        return {
            "found_keywords": found,
            "has_pdf": has_pdf,
            "title": title,
        }

    async def run(self):
        print("=" * 60)
        print("The Royal United Services Institute 文章爬蟲")
        print("=" * 60)

        pw = await async_playwright().start()
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage"]
        )
        context = await browser.new_context(
            viewport={"width": 1280, "height": 800},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        )
        page = await context.new_page()

        try:
            for search_kw in self.KEYWORDS:
                display_kw = self.KEYWORD_DISPLAY[search_kw]
                print(f"\n{'-'*20}")
                print(f"🔍 搜尋關鍵字：{display_kw}")
                print("-"*20)

                await page.goto(f"{self.BASE_URL}/search", timeout=60000, wait_until="networkidle")
                await asyncio.sleep(2)

                js_fill = (
                    "(kw) => {"
                    "const inputs = Array.from(document.querySelectorAll('input'));"
                    "const box = inputs.find(el => el.offsetParent !== null);"
                    "if (!box) return false;"
                    "box.focus();"
                    "const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;"
                    "setter.call(box, kw);"
                    "box.dispatchEvent(new Event('input', {bubbles:true}));"
                    "box.dispatchEvent(new Event('change', {bubbles:true}));"
                    "return true;}"
                )
                await page.evaluate(js_fill, display_kw)
                await asyncio.sleep(1)
                await page.keyboard.press("Enter")
                await page.wait_for_load_state("networkidle", timeout=20000)
                await asyncio.sleep(1)

                print(f"  取得文章清單中...")
                candidates = self.get_articles_in_range(search_kw)
                print(f"  符合日期範圍：{len(candidates)} 篇，開始逐篇檢查...\n")

                for i, art in enumerate(candidates, 1):
                    url   = art["url"]
                    adate = art["date"]
                    date_str = adate.strftime("%Y/%m/%d") if adate else "無日期"

                    if url in self.visited_urls:
                        print(f"  [{i}/{len(candidates)}] ⏭️  已拜訪過，跳過  ({date_str})")
                        continue
                    self.visited_urls.add(url)

                    print(f"  [{i}/{len(candidates)}] {date_str}  {art['title'][:55]}")

                    try:
                        await page.goto(url, timeout=20000, wait_until="domcontentloaded")
                        await asyncio.sleep(0.5)
                    except Exception:
                        pass

                    result = self.check_article(url, self.KEYWORDS)
                    if result is None:
                        continue

                    if result.get("skip_reason") == "members_only":
                        print(f"         🔒 Members Only，跳過")
                        continue

                    kws     = result["found_keywords"]
                    has_pdf = result["has_pdf"]
                    title   = result["title"] or art["title"]

                    pdf_tag = "✅ PDF" if has_pdf else "❌ 無PDF"
                    kw_tag  = f"關鍵字: {', '.join(kws)}" if kws else "❌ 無關鍵字"
                    print(f"         {pdf_tag}  {kw_tag}")

                    if kws and has_pdf:
                        self.results.append({
                            "機構": self.INSTITUTION,
                            "日期": date_str,
                            "標題": title,
                            "網址": url,
                            "關鍵字": ", ".join(kws),
                        })
                        print(f"         ✅ 納入！目前共 {len(self.results)} 篇")

                    time.sleep(0.3)

            await page.goto(f"{self.BASE_URL}/search", timeout=15000)

        finally:
            print(f"🎉 RUSI 完成！共找到 {len(self.results)} 篇符合文章")
            await browser.close()
            await pw.stop()

        return self.results


# ═══════════════════════════════════════════════════════════════════════
# MERICS Scraper Class
# ═══════════════════════════════════════════════════════════════════════

class MericsScraper:

    def __init__(self, start_dt, end_dt, visited_urls):
        self.start_date = start_dt
        self.end_date = end_dt
        self.visited_urls = visited_urls
        self.results = []

        self.excluded_url_patterns = [
            '/en/germany-china', '/en/us-china', '/en/russia-china',
            '/en/geopolitics', '/en/trade-and-investment',
            '/en/industrial-policy-and-technology', '/en/party-and-state',
            '/en/digital-china', '/en/belt-and-road', '/en/hong-kong',
            '/en/climate-and-environment', '/en/chinese-debates', '/en/podcast',
            '/en/media', '/en/events', '/en/about', '/en/contact', '/en/experts',
            '/en/leadership-and-staff', '/en/services', '/en/opportunities',
            '/en/governance', '/en/partners', '/en/become-a-member',
            '/en/merics-update', '/en/merics-eu-china-hub',
            '/en/privacy-policy-and-cookies', '/en/legal-notice', '/en/rss',
            '/en/user/login', '/en/node/', '/en/search-results'
        ]

        self.driver = webdriver.Chrome(options=build_chrome_options())
        self.wait = WebDriverWait(self.driver, 10)

    def parse_date(self, date_str):
        try:
            return datetime.strptime(date_str.strip(), "%b %d, %Y")
        except:
            try:
                return datetime.strptime(date_str.strip(), "%B %d, %Y")
            except:
                print(f"    無法解析日期: {date_str}")
                return None

    def check_article_details(self, article_url, keyword):
        try:
            for pattern in self.excluded_url_patterns:
                if pattern in article_url:
                    print(f"      → 跳過分類頁面: {article_url}")
                    return None

            print(f"      → 檢查文章內頁: {article_url}")

            main_window = self.driver.current_window_handle
            self.driver.execute_script(f"window.open('{article_url}', '_blank');")
            self.driver.switch_to.window(self.driver.window_handles[-1])
            time.sleep(2)

            title_text = None
            title_selectors = [
                "//h1", "//h1[@class='title']",
                "//*[contains(@class, 'page-title')]",
                "//*[contains(@class, 'article-title')]"
            ]
            for selector in title_selectors:
                try:
                    title_elements = self.driver.find_elements(By.XPATH, selector)
                    if title_elements:
                        title_text = title_elements[0].text.strip()
                        if title_text:
                            break
                except:
                    continue

            if not title_text:
                print(f"      ✗ 找不到標題")
                self.driver.close()
                self.driver.switch_to.window(main_window)
                return None

            print(f"      標題: {title_text[:60]}...")

            date_str = None
            date_selectors = [
                "//time", "//*[contains(@class, 'date')]",
                "//*[contains(@class, 'published')]",
                "//*[contains(text(), '202')]"
            ]
            for selector in date_selectors:
                try:
                    date_elements = self.driver.find_elements(By.XPATH, selector)
                    if date_elements:
                        date_str = date_elements[0].text.strip()
                        if date_str and ('2025' in date_str or '2026' in date_str):
                            break
                except:
                    continue

            if not date_str:
                print(f"      ✗ 找不到日期")
                self.driver.close()
                self.driver.switch_to.window(main_window)
                return None

            article_date = self.parse_date(date_str)
            if not article_date:
                print(f"      ✗ 日期解析失敗 ({date_str})")
                self.driver.close()
                self.driver.switch_to.window(main_window)
                return None

            print(f"      日期: {article_date.strftime('%Y/%m/%d')}")

            if article_date < self.start_date:
                print(f"      ⚠ 早於起始日期 ({self.start_date.strftime('%Y/%m/%d')})")
                self.driver.close()
                self.driver.switch_to.window(main_window)
                return "TOO_OLD"

            if article_date > self.end_date:
                print(f"      ✗ 晚於結束日期")
                self.driver.close()
                self.driver.switch_to.window(main_window)
                return None

            def _looks_like_pdf_url(url: str) -> bool:
                if not url:
                    return False
                u = url.lower().split("#")[0]
                return u.endswith(".pdf") or ".pdf?" in u

            def _head_is_pdf(url: str, timeout=10) -> bool:
                try:
                    r = req_lib.head(url, allow_redirects=True, timeout=timeout)
                    ct = (r.headers.get("Content-Type") or "").lower()
                    if "application/pdf" in ct:
                        return True
                    r = req_lib.get(url, stream=True, allow_redirects=True, timeout=timeout)
                    ct = (r.headers.get("Content-Type") or "").lower()
                    if "application/pdf" in ct:
                        return True
                    cd = (r.headers.get("Content-Disposition") or "").lower()
                    if ".pdf" in cd:
                        return True
                except Exception:
                    pass
                return False

            has_download = False

            pdf_link_xpaths = [
                "//*[self::a or self::button][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]",
                "//*[self::a or self::button][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pdf version')]",
                "//*[self::a or self::button][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'full pdf')]",
                "//*[self::a or self::button][contains(normalize-space(.), '(pdf')]",
                "//*[self::a or self::button][contains(normalize-space(.), '.pdf')]",
            ]

            download_like_elements = []
            seen_elements = set()
            for xpath in pdf_link_xpaths:
                try:
                    elements = self.driver.find_elements(By.XPATH, xpath)
                    for el in elements:
                        element_id = id(el)
                        if element_id not in seen_elements:
                            download_like_elements.append(el)
                            seen_elements.add(element_id)
                except:
                    continue

            for el in download_like_elements:
                try:
                    text = (el.text or "").strip().lower()
                    href = (el.get_attribute("href") or "").strip()
                    if "print" in text or "share" in text:
                        continue
                    if not ("download" in text or "pdf" in text or ".pdf" in text):
                        continue
                    if not href or not _looks_like_pdf_url(href):
                        continue
                    if _head_is_pdf(href):
                        has_download = True
                        print(f"      找到 PDF 連結: {text[:50]}...")
                        break
                except Exception:
                    continue

            self.driver.close()
            self.driver.switch_to.window(main_window)

            if not has_download:
                print(f"      ✗ 沒有 DOWNLOAD 按鈕")
                return None

            print(f"      ✓ 有 DOWNLOAD 按鈕")

            return {
                '機構': 'Mercator Institute for China Studies',
                '日期': article_date.strftime('%Y/%m/%d'),
                '標題': title_text,
                '網址': article_url,
                '關鍵字': keyword
            }

        except Exception as e:
            print(f"      ✗ 檢查文章時發生錯誤: {str(e)}")
            try:
                self.driver.close()
                self.driver.switch_to.window(main_window)
            except:
                pass
            return None

    def scrape_page(self, url, keyword):
        page_num = 0
        should_continue = True
        consecutive_old_articles = 0
        processed_urls = set()

        while should_continue:
            try:
                current_url = url if page_num == 0 else f"{url}&page={page_num}"
                print(f"\n  正在爬取頁面 {page_num}: {current_url}")
                self.driver.get(current_url)
                time.sleep(3)

                allowed_prefix = None
                if "field_publication_type=4" in url:
                    allowed_prefix = "https://merics.org/en/report/"
                elif "field_publication_type=15" in url:
                    allowed_prefix = "https://merics.org/en/merics-briefs/"
                elif "field_publication_type=13" in url:
                    allowed_prefix = "https://merics.org/en/comment/"
                elif "field_publication_type=9" in url:
                    allowed_prefix = "https://merics.org/en/tracker/"

                article_links = []
                seen_links = set()

                selectors = [
                    "//h2//a[@href]", "//h3//a[@href]",
                    "//article//a[@href]",
                    "//*[contains(@class, 'views-row')]//a[@href]"
                ]

                for selector in selectors:
                    try:
                        links = self.driver.find_elements(By.XPATH, selector)
                        for link in links:
                            href = link.get_attribute('href')
                            if href and href not in seen_links:
                                if '/en/' in href and 'merics.org' in href:
                                    if allowed_prefix and not href.startswith(allowed_prefix):
                                        continue
                                    article_links.append(href)
                                    seen_links.add(href)
                    except Exception as e:
                        continue

                if not article_links:
                    print(f"    頁面 {page_num} 沒有找到文章連結")
                    break

                filtered_links = []
                for href in article_links:
                    is_excluded = False
                    for pattern in self.excluded_url_patterns:
                        if pattern in href:
                            is_excluded = True
                            break
                    if not is_excluded:
                        filtered_links.append(href)

                article_links = filtered_links
                new_links = [link for link in article_links if link not in processed_urls and link not in self.visited_urls]

                if not new_links:
                    print(f"    頁面 {page_num} 沒有新的文章連結（已全部處理過）")
                    if page_num < 3:
                        page_num += 1
                        continue
                    else:
                        print("  已嘗試多個頁面，沒有更多新文章，停止爬取")
                        break

                print(f"    找到 {len(article_links)} 個連結，其中 {len(new_links)} 個是新文章")

                for idx, article_url in enumerate(new_links):
                    print(f"\n    處理文章 {idx + 1}/{len(new_links)}:")
                    print(f"      URL: {article_url}")

                    processed_urls.add(article_url)
                    self.visited_urls.add(article_url)

                    article_info = self.check_article_details(article_url, keyword)

                    if article_info == "TOO_OLD":
                        consecutive_old_articles += 1
                        print(f"      ⚠ 連續過舊文章計數: {consecutive_old_articles}/3")
                        if consecutive_old_articles >= 3:
                            print(f"\n    ⚠⚠⚠ 連續 3 篇過舊文章，停止爬取")
                            should_continue = False
                            break
                    elif article_info:
                        consecutive_old_articles = 0
                        self.results.append(article_info)
                        print(f"      ✓✓✓ 收錄成功！")
                    else:
                        print(f"      ✗ 不符合條件（不影響過舊文章計數）")

                if not should_continue:
                    break

                page_num += 1
                print(f"\n  準備載入頁面 {page_num}...")
                time.sleep(2)

            except Exception as e:
                print(f"  處理頁面時發生錯誤: {str(e)}")
                break

    def scrape_all(self):
        base_urls = [
            "https://merics.org/en/analysis?search_api_fulltext={}&field_topic=All&field_publication_type=4&sort_by=field_date_published",
            "https://merics.org/en/analysis?search_api_fulltext={}&field_topic=All&field_publication_type=15&sort_by=field_date_published",
            "https://merics.org/en/analysis?search_api_fulltext={}&field_topic=All&field_publication_type=13&sort_by=field_date_published",
            "https://merics.org/en/analysis?search_api_fulltext={}&field_topic=All&field_publication_type=9&sort_by=field_date_published"
        ]
        keywords = ["china", "taiwan", "taipei"]

        for keyword in keywords:
            keyword_capitalized = keyword.capitalize()
            print(f"\n{'-' * 20}")
            print(f"開始爬取關鍵字: {keyword.upper()}")
            print(f"{'-' * 20}")

            for base_url in base_urls:
                url = base_url.format(keyword)
                try:
                    self.scrape_page(url, keyword_capitalized)
                except Exception as e:
                    print(f"  爬取時發生錯誤: {str(e)}")
                    continue

    def run(self):
        print("=" * 60)
        print("MERICS 網站爬蟲")
        print("=" * 60)

        try:
            self.scrape_all()
            print(f"MERICS 爬取完成！共收錄 {len(self.results)} 篇文章")
        except Exception as e:
            print(f"\n發生錯誤: {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            print("\n關閉瀏覽器...")
            self.driver.quit()
            print("✓ MERICS 瀏覽器已關閉")

        return self.results


# ═══════════════════════════════════════════════════════════════════════
# RAND Scraper Class
# ═══════════════════════════════════════════════════════════════════════

class RANDScraper:

    def __init__(self, start_dt, end_dt, visited_urls):
        self.start_date = start_dt
        self.end_date = end_dt
        self.visited_urls = visited_urls
        self.driver = None
        self.results = []

    def setup_driver(self):
        self.driver = webdriver.Chrome(options=build_chrome_options())

    def parse_article_date(self, date_text):
        if not date_text:
            return None

        months = {
            'January': 1, 'February': 2, 'March': 3, 'April': 4,
            'May': 5, 'June': 6, 'July': 7, 'August': 8,
            'September': 9, 'October': 10, 'November': 11, 'December': 12,
            'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
            'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
        }

        pattern = r'(\w+)\s+(\d{1,2}),?\s+(\d{4})'
        match = re.search(pattern, date_text)

        if match:
            month_str, day_str, year_str = match.groups()
            if month_str in months:
                try:
                    return datetime(int(year_str), months[month_str], int(day_str))
                except:
                    return None

        return None

    def check_download_button(self):
        try:
            all_links_buttons = self.driver.find_elements(By.XPATH, "//a | //button")
            for e in all_links_buttons:
                if e.is_displayed() and 'download' in e.text.strip().lower():
                    return True
            return False
        except:
            return False

    def scrape_keyword(self, keyword):
        base_url = (
            f"https://www.rand.org/search.html?q={keyword.lower()}"
            f"&content_type_ss=Research&content_type_ss=Commentary"
            f"&content_type_ss=Article&content_type_ss=Research+Summary"
            f"&dateFixedRange=Last+30+days&sortby=date_dt&rows=12"
        )

        print(f"\n🔍 正在搜尋關鍵字: {keyword}")
        self.driver.get(base_url)
        time.sleep(5)

        try:
            accept_button = self.driver.find_element(By.XPATH,
                "//*[contains(text(), 'Accept') or contains(text(), 'accept')]")
            accept_button.click()
            time.sleep(1)
        except:
            pass

        consecutive_early_count = 0
        page = 1

        while consecutive_early_count < 3:
            print(f"  📄 第 {page} 頁")

            urls = []
            try:
                time.sleep(2)

                selectors = [
                    "h3.search-result-title a", "h3 a", ".search-result a",
                    "a[href*='/pubs/']", "div.result a", "article a"
                ]

                for selector in selectors:
                    article_links = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if article_links:
                        urls = [link.get_attribute('href') for link in article_links
                                if link.get_attribute('href') and '/pubs/' in link.get_attribute('href')]
                        if urls:
                            print(f"  ✅ 找到 {len(urls)} 篇文章")
                            break

            except Exception as e:
                print(f"  ⚠️  找不到文章連結: {str(e)}")

            if not urls:
                print("  ⚠️  沒有更多文章")
                break

            for url in urls:
                try:
                    if url in self.visited_urls:
                        print(f"    ⏭️  已訪問，略過: {url[:60]}")
                        continue

                    self.driver.get(url)
                    time.sleep(2)

                    try:
                        title = self.driver.find_element(By.CSS_SELECTOR, "h1").text.strip()
                    except:
                        title = "無標題"

                    article_date = None
                    date_str_display = None

                    try:
                        time.sleep(1)
                        possible_date_elements = self.driver.find_elements(By.XPATH,
                            "//*[contains(text(), '2024') or contains(text(), '2025') or contains(text(), '2026')]")

                        for elem in possible_date_elements:
                            text = elem.text.strip()
                            if len(text) < 50:
                                parsed_date = self.parse_article_date(text)
                                if parsed_date:
                                    article_date = parsed_date
                                    date_str_display = article_date.strftime('%Y/%m/%d')
                                    break

                        if not article_date:
                            date_elements = self.driver.find_elements(By.CSS_SELECTOR,
                                ".date, .published, .pub-date, time, [class*='date']")
                            for elem in date_elements:
                                text = elem.text.strip()
                                if text:
                                    parsed_date = self.parse_article_date(text)
                                    if parsed_date:
                                        article_date = parsed_date
                                        date_str_display = article_date.strftime('%Y/%m/%d')
                                        break

                    except Exception as e:
                        print(f"      ⚠️  解析日期時發生錯誤: {str(e)}")

                    if not article_date:
                        print(f"    ⏭️  找不到發布日期: {title[:50]}...")
                        self.visited_urls.add(url)
                        continue

                    if article_date < self.start_date:
                        consecutive_early_count += 1
                        print(f"    ⏭️  發布日期早於起始時間 ({consecutive_early_count}/3): {title[:50]}... ({date_str_display})")
                        self.visited_urls.add(url)
                        if consecutive_early_count >= 3:
                            print(f"  ✅ 找到連續 3 篇早於起始時間的文章,停止搜尋")
                            break
                        continue

                    consecutive_early_count = 0

                    if article_date > self.end_date:
                        print(f"    ⏭️  發布日期晚於結束時間: {title[:50]}... ({date_str_display})")
                        self.visited_urls.add(url)
                        continue

                    has_download = self.check_download_button()
                    self.visited_urls.add(url)

                    if has_download:
                        self.results.append({
                            '機構': 'RAND',
                            '日期': date_str_display,
                            '標題': title,
                            '網址': url,
                            '關鍵字': keyword.capitalize()
                        })
                        print(f"    ✅ 收錄: {title[:50]}... ({date_str_display})")
                    else:
                        print(f"    ⏭️  無 Download 按鍵: {title[:50]}... ({date_str_display})")

                except Exception as e:
                    print(f"    ❌ 處理文章時發生錯誤: {str(e)}")
                    continue

            if consecutive_early_count >= 3:
                break

            try:
                self.driver.back()
                time.sleep(2)
                next_buttons = self.driver.find_elements(By.XPATH,
                    "//a[contains(text(), 'Next') or contains(@aria-label, 'Next')]")
                if next_buttons:
                    next_buttons[0].click()
                    time.sleep(3)
                    page += 1
                else:
                    print("  ⚠️  沒有下一頁")
                    break
            except:
                print("  ⚠️  翻頁失敗")
                break

    def run(self):
        print("=" * 60)
        print("RAND Corporation 網站爬蟲")
        print("=" * 60)

        try:
            print("\n🚀 啟動瀏覽器...")
            self.setup_driver()

            keywords = ['china', 'taiwan', 'taipei']
            for keyword in keywords:
                self.scrape_keyword(keyword)

            print(f"RAND 爬取完成！共收錄 {len(self.results)} 篇文章")

        except Exception as e:
            print(f"\n❌ 發生錯誤: {str(e)}")

        finally:
            if self.driver:
                print("\n關閉瀏覽器...")
                self.driver.quit()
                print("✓ RAND 瀏覽器已關閉")

        return self.results


# ═══════════════════════════════════════════════════════════════════════
# JISS Scraper Class
# ═══════════════════════════════════════════════════════════════════════

class JISSScraper:

    INSTITUTION = "Jerusalem Institute for Strategy and Security"
    KEYWORDS    = ["China", "Taiwan", "Taipei"]
    MONTH_MAP   = {
        "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
        "july":7,"august":8,"september":9,"october":10,"november":11,"december":12,
        "jan":1,"feb":2,"mar":3,"apr":4,"jun":6,"jul":7,"aug":8,
        "sep":9,"oct":10,"nov":11,"dec":12
    }

    def __init__(self, start_dt, end_dt, visited_urls):
        self.start_dt = start_dt
        self.end_dt = end_dt
        self.visited_urls = visited_urls
        self.results = []
        self.driver = None

    def setup_browser(self):
        self.driver = webdriver.Chrome(options=build_chrome_options())
        self.driver.set_page_load_timeout(30)

    def parse_date(self, text):
        text = text.strip()
        m = re.search(r'(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})', text)
        if m:
            day, mon, year = int(m.group(1)), m.group(2).lower(), int(m.group(3))
            if mon in self.MONTH_MAP:
                return datetime(year, self.MONTH_MAP[mon], day)
        m = re.search(r'([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})', text)
        if m:
            mon, day, year = m.group(1).lower(), int(m.group(2)), int(m.group(3))
            if mon in self.MONTH_MAP:
                return datetime(year, self.MONTH_MAP[mon], day)
        m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', text)
        if m:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', text)
        if m:
            day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1 <= month <= 12 and 1 <= day <= 31:
                return datetime(year, month, day)
        return None

    def has_printfriendly_icon(self):
        try:
            imgs = self.driver.find_elements(By.TAG_NAME, "img")
            for img in imgs:
                src  = img.get_attribute("src")  or ""
                dsrc = img.get_attribute("data-lazy-src") or ""
                cls  = img.get_attribute("class") or ""
                if "printfriendly" in src.lower() or "printfriendly" in dsrc.lower():
                    return True
                if "pf-button-img" in cls:
                    return True
            pf = self.driver.find_elements(By.CSS_SELECTOR, ".pf-button-img, [class*='pf-button']")
            return len(pf) > 0
        except Exception:
            return False

    def get_pdf_link(self):
        try:
            links = self.driver.find_elements(By.TAG_NAME, "a")
            for lnk in links:
                text = lnk.text.strip().lower()
                if "click here to read" in text or "click here to download" in text:
                    href = lnk.get_attribute("href") or ""
                    if ".pdf" in href.lower():
                        return href
        except Exception:
            pass
        return None

    def find_keywords(self):
        found = []
        try:
            title_text = ""
            for sel in ["h1.entry-title", "h1.post-title", "h1", ".entry-title", ".post-title"]:
                els = self.driver.find_elements(By.CSS_SELECTOR, sel)
                if els:
                    title_text = els[0].text
                    break

            content_text = ""
            for sel in [
                ".entry-content", ".post-content", "article .content",
                ".elementor-widget-theme-post-content", "article", ".single-post-content"
            ]:
                els = self.driver.find_elements(By.CSS_SELECTOR, sel)
                if els:
                    content_text = els[0].text
                    break

            content_clean = content_text
            for marker in [
                "\nReferences\n", "\nNotes\n", "\nBibliography\n", "\nEndnotes\n",
                "\nreferences\n", "\nnotes\n", "\nSources\n", "\nSources:\n", "\nsources\n",
            ]:
                idx = content_clean.find(marker)
                if idx != -1:
                    content_clean = content_clean[:idx]
                    break

            lines = content_clean.splitlines()
            clean_lines = []
            for line in lines:
                stripped = line.strip()
                if re.match(r'^\[(?:[ivxlcdmIVXLCDM]+|\d+)\]', stripped):
                    break
                clean_lines.append(line)
            content_clean = "\n".join(clean_lines)

            combined = (title_text + " " + content_clean).lower()
            for kw in self.KEYWORDS:
                if kw.lower() in combined:
                    found.append(kw)
        except Exception:
            pass
        return found

    def extract_card_date(self, card):
        for date_sel in [
            ".post-date", ".entry-date", "time", ".date", "[class*='date']",
            ".jet-listing-dynamic-field", ".elementor-post__meta-data"
        ]:
            try:
                el = card.find_element(By.CSS_SELECTOR, date_sel)
                dt_attr = el.get_attribute("datetime")
                date_text = dt_attr if dt_attr else el.text
                card_date = self.parse_date(date_text)
                if card_date:
                    return card_date
            except NoSuchElementException:
                continue
        try:
            return self.parse_date(card.text)
        except Exception:
            return None

    def extract_article_url(self, card, base_url):
        for link_sel in [
            "h2 a", "h3 a", "h4 a",
            ".entry-title a", ".post-title a", ".jet-listing-dynamic-field a",
            ".elementor-post__title a", ".elementor-heading-title a",
        ]:
            try:
                link_el = card.find_element(By.CSS_SELECTOR, link_sel)
                href = link_el.get_attribute("href") or ""
                if href and "jiss.org.il" in href and "/author/" not in href:
                    return href
            except NoSuchElementException:
                continue
        try:
            all_links = card.find_elements(By.CSS_SELECTOR, "a[href]")
            for lnk in all_links:
                href = lnk.get_attribute("href") or ""
                if (href and "jiss.org.il" in href
                        and "/author/" not in href
                        and "/tag/" not in href
                        and "/category/" not in href
                        and base_url.rstrip("/") not in href.rstrip("/")):
                    return href
        except Exception:
            pass
        return None

    def open_tab(self, url):
        original_window = self.driver.current_window_handle
        self.driver.execute_script("window.open(arguments[0], '_blank');", url)
        time.sleep(1)
        self.driver.switch_to.window(self.driver.window_handles[-1])
        try:
            WebDriverWait(self.driver, 25).until(
                lambda d: d.execute_script("return document.readyState") in ("complete", "interactive")
            )
            return True, original_window
        except Exception:
            print(f"        ⚠️ 頁面載入超時，跳過")
            self.driver.close()
            self.driver.switch_to.window(original_window)
            time.sleep(1)
            return False, original_window

    def scrape_section(self, base_url, mode):
        page_num = 1
        stop_scraping = False

        while not stop_scraping:
            url = base_url if page_num == 1 else f"{base_url}{page_num}/"
            print(f"  📄 正在載入第 {page_num} 頁：{url}")
            try:
                self.driver.get(url)
            except Exception:
                self.driver.execute_script("window.stop();")
            time.sleep(2)

            article_cards = []
            for sel in [
                "article.type-post", ".post-item", ".articles-list article",
                ".elementor-post", ".jet-listing-grid__item",
                "article", ".post"
            ]:
                cards = self.driver.find_elements(By.CSS_SELECTOR, sel)
                if cards:
                    article_cards = cards
                    break

            if not article_cards:
                print("  ⚠ 找不到文章列表，可能已到末頁，停止。")
                break

            print(f"     找到 {len(article_cards)} 篇文章")
            consecutive_early = 0

            for card in article_cards:
                card_date = self.extract_card_date(card)
                print(f"  📅 日期：{card_date.strftime('%Y/%m/%d') if card_date else '（無法解析）'}")

                if card_date:
                    if card_date > self.end_dt:
                        continue
                    if card_date < self.start_dt:
                        consecutive_early += 1
                        print(f"     ⏩ 早於起始日期 ({card_date.strftime('%Y/%m/%d')})，跳過 [{consecutive_early}/3]")
                        if consecutive_early >= 3:
                            print("     🛑 連續 3 篇早於起始日期，停止爬取。")
                            stop_scraping = True
                            break
                        continue
                    else:
                        consecutive_early = 0

                article_url = self.extract_article_url(card, base_url)
                if not article_url:
                    continue

                if article_url in self.visited_urls:
                    print(f"     ⏭️ 已訪問，略過：{article_url[:60]}")
                    continue

                print(f"     🔍 開啟文章：{article_url}")
                success, original_window = self.open_tab(article_url)
                if not success:
                    continue

                try:
                    found_kws = self.find_keywords()

                    if mode == "printfriendly":
                        condition_met = self.has_printfriendly_icon()
                        condition_label = f"PrintFriendly Icon: {'✔' if condition_met else '✗'}"
                        record_url = article_url
                    else:
                        pdf_url = self.get_pdf_link()
                        condition_met = pdf_url is not None
                        condition_label = f"PDF按鈕: {'✔' if condition_met else '✗'}"
                        record_url = pdf_url if pdf_url else article_url

                    print(f"        關鍵字: {found_kws or '無'}  |  {condition_label}")

                    self.visited_urls.add(article_url)

                    if found_kws and condition_met:
                        title = self.driver.title.replace(" - JISS", "").replace(" | JISS", "").strip()
                        for sel in ["h1.entry-title", "h1.post-title", "h1", ".entry-title"]:
                            els = self.driver.find_elements(By.CSS_SELECTOR, sel)
                            if els:
                                title = els[0].text.strip()
                                break

                        if card_date:
                            date_str = card_date.strftime("%Y/%m/%d")
                        else:
                            date_str = ""
                            for dsel in [".post-date", ".entry-date", "time", ".date"]:
                                try:
                                    el = self.driver.find_element(By.CSS_SELECTOR, dsel)
                                    dt_attr = el.get_attribute("datetime")
                                    txt = dt_attr if dt_attr else el.text
                                    parsed = self.parse_date(txt)
                                    if parsed:
                                        date_str = parsed.strftime("%Y/%m/%d")
                                        break
                                except Exception:
                                    continue

                        self.results.append({
                            "機構": self.INSTITUTION,
                            "日期": date_str,
                            "標題": title,
                            "網址": record_url,
                            "關鍵字": ", ".join(found_kws)
                        })
                        print(f"        ✅ 收錄：{title}")
                finally:
                    self.driver.close()
                    self.driver.switch_to.window(original_window)
                    time.sleep(1)

            if stop_scraping:
                break

            page_num += 1
            next_url = f"{base_url}{page_num}/"
            try:
                self.driver.get(next_url)
            except Exception:
                self.driver.execute_script("window.stop();")
            time.sleep(2)
            current = self.driver.current_url.rstrip("/")
            expected = next_url.rstrip("/")
            section_slug = base_url.split("/en/")[1].strip("/")
            if current != expected and section_slug not in current:
                print("  ✅ 已到最後一頁，結束爬取。")
                break
            test_cards = []
            for sel in ["article.type-post", ".post-item", "article", ".post"]:
                test_cards = self.driver.find_elements(By.CSS_SELECTOR, sel)
                if test_cards:
                    break
            if not test_cards:
                print("  ✅ 下一頁無文章，結束爬取。")
                break

    def run(self):
        print("=" * 60)
        print("Jerusalem Institute for Strategy and Security文章爬蟲")
        print("=" * 60)

        self.setup_browser()

        try:
            print("-" * 20)
            print("Jerusalem Papers")
            print("-" * 20)
            self.scrape_section("https://jiss.org.il/en/jerusalem-papers/", mode="pdf_button")

            print("\n" + "-" * 20)
            print("In-Depth Analysis")
            print("-" * 20)
            self.scrape_section("https://jiss.org.il/en/in-depth-analysis/", mode="pdf_button")

            print("\n" + "-" * 20)
            print("Articles")
            print("-" * 20)
            self.scrape_section("https://jiss.org.il/en/articles/", mode="printfriendly")

            print(f"\n  共收錄 {len(self.results)} 篇文章")

        except Exception as e:
            print(f"\n❌ 執行過程中發生錯誤: {e}")
            import traceback
            traceback.print_exc()
        finally:
            self.driver.quit()
            print("✓ JISS 瀏覽器已關閉")

        return self.results


# ═══════════════════════════════════════════════════════════════════════
# 匯出 Excel
# ═══════════════════════════════════════════════════════════════════════

def export_excel(rows: list, output_path: str = "combined_articles_G1_normal.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Articles"
    ws.append(["機構", "日期", "標題", "網址", "關鍵字"])
    for row in rows:
        ws.append([row["機構"], row["日期"], row["標題"], row["網址"], row["關鍵字"]])
    wb.save(output_path)
    print(f"✅ 已匯出 {output_path}（共 {len(rows)} 筆）")


# ═══════════════════════════════════════════════════════════════════════
# 主程式
# ═══════════════════════════════════════════════════════════════════════

async def main():

    # ── AC Global China Hub ───────────────────────────────────────────
    try:
        gch_scraper = ACGlobalChinaHubScraper(start_dt, end_dt, visited_urls)
        gch_results = gch_scraper.run()
        all_results.extend(gch_results)
    except Exception as e:
        print(f"\n⚠️  AC Global China Hub 爬取發生錯誤，跳過：{e}")

    # ── Atlantic Council ──────────────────────────────────────────────
    try:
        ac_scraper = AtlanticCouncilScraper(start_dt, end_dt, visited_urls)
        ac_results = ac_scraper.run()
        all_results.extend(ac_results)
    except Exception as e:
        print(f"\n⚠️  Atlantic Council 爬取發生錯誤，跳過：{e}")

    # ── RUSI ──────────────────────────────────────────────────────────
    try:
        rusi_scraper = RUSIScraper(start_dt, end_dt, visited_urls)
        rusi_results = await rusi_scraper.run()
        all_results.extend(rusi_results)
    except Exception as e:
        print(f"\n⚠️  RUSI 爬取發生錯誤，跳過：{e}")

    # ── MERICS ────────────────────────────────────────────────────────
    try:
        merics_scraper = MericsScraper(start_dt, end_dt, visited_urls)
        merics_results = merics_scraper.run()
        all_results.extend(merics_results)
    except Exception as e:
        print(f"\n⚠️  MERICS 爬取發生錯誤，跳過：{e}")

    # ── RAND ──────────────────────────────────────────────────────────
    try:
        rand_scraper = RANDScraper(start_dt, end_dt, visited_urls)
        rand_results = rand_scraper.run()
        all_results.extend(rand_results)
    except Exception as e:
        print(f"\n⚠️  RAND 爬取發生錯誤，跳過：{e}")

    # ── JISS ──────────────────────────────────────────────────────────
    try:
        jiss_scraper = JISSScraper(start_dt, end_dt, visited_urls)
        jiss_results = jiss_scraper.run()
        all_results.extend(jiss_results)
    except Exception as e:
        print(f"\n⚠️  JISS 爬取發生錯誤，跳過：{e}")

    # ── 匯出 Excel ────────────────────────────────────────────────────
    excel_path = "combined_articles_G1_normal.xlsx"
    export_excel(all_results, excel_path)

    # ── 寄送 Email（含 Excel + log）──────────────────────────────────
    log_text = logger.get_log()
    send_email(excel_path, log_text)


if __name__ == "__main__":
    asyncio.run(main())
