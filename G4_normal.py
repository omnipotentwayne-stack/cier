# -*- coding: utf-8 -*-
import io
import sys
import os
import re
import time
import calendar
from datetime import datetime, timedelta
from typing import List, Dict, Tuple
from io import BytesIO

import requests
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import PyPDF2
    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False

try:
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False


# ═══════════════════════════════════════════════════════════════════════
# LOG 攔截
# ═══════════════════════════════════════════════════════════════════════

class TeeLogger:
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
# 自動計算日期區間
# ═══════════════════════════════════════════════════════════════════════

today = datetime.today()
day = today.day

if day == 1:
    first_of_this_month = today.replace(day=1)
    last_month = first_of_this_month - timedelta(days=1)
    last_day = calendar.monthrange(last_month.year, last_month.month)[1]
    start_date = last_month.replace(day=21)
    end_date   = last_month.replace(day=last_day)
elif day == 11:
    start_date = today.replace(day=1)
    end_date   = today.replace(day=10)
elif day == 21:
    start_date = today.replace(day=11)
    end_date   = today.replace(day=20)
else:
    end_date   = today
    start_date = today - timedelta(days=10)
    print(f"  ⚠ 今天是 {day} 號，非標準排程日，使用預設區間（最近10天）")

start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
end_date   = end_date.replace(hour=23, minute=59, second=59, microsecond=0)

print(f"  ✓ 時間區間: {start_date.date()} ~ {end_date.date()}")


# ═══════════════════════════════════════════════════════════════════════
# 寄信功能
# ═══════════════════════════════════════════════════════════════════════

def send_email(excel_path: str, log_text: str, total_count: int):
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
    msg["Subject"] = f"G4爬蟲結果 {today.strftime('%Y/%m/%d')}（區間 {start_date.date()} ~ {end_date.date()}）"

    body = (
        f"G4爬蟲執行完畢。\n"
        f"執行日期：{today.strftime('%Y/%m/%d %H:%M')}\n"
        f"資料區間：{start_date.date()} ~ {end_date.date()}\n"
        f"共收錄：{total_count} 筆\n\n"
        f"詳細 log 請見附件 G4_log.txt，Excel 結果請見附件。"
    )
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename={os.path.basename(excel_path)}")
            msg.attach(part)

    log_part = MIMEBase("text", "plain")
    log_part.set_payload(log_text.encode("utf-8"))
    encoders.encode_base64(log_part)
    log_part.add_header("Content-Disposition", "attachment; filename=G4_log.txt")
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
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    return opts


def make_driver():
    opts = build_chrome_options()
    if USE_WEBDRIVER_MANAGER:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=opts)
    return webdriver.Chrome(options=opts)


# ════════════════════════════════════════════════════════════
#  GatewayHouse Scraper
# ════════════════════════════════════════════════════════════

class GatewayHouseScraper:
    URLS = [
        "https://www.gatewayhouse.in/publications/research-reports/",
        "https://www.gatewayhouse.in/publications/articles/",
        "https://www.gatewayhouse.in/publications/conference-papers/",
        "https://www.gatewayhouse.in/publications/research-papers/",
        "https://www.gatewayhouse.in/publications/books/"
    ]
    KEYWORDS    = ["China", "Taiwan", "Taipei"]
    INSTITUTION = "Indian Council on Global Relations"

    def __init__(self, start_date, end_date):
        self.driver     = make_driver()
        self.articles   = []
        self.start_date = start_date
        self.end_date   = end_date

    def parse_article_date(self, date_text):
        try:
            date_text = date_text.strip().replace(',', '')
            return datetime.strptime(date_text, "%d %B %Y")
        except:
            try:
                return datetime.strptime(date_text, "%B %d %Y")
            except:
                return None

    def find_keywords_in_text(self, text):
        return [kw for kw in self.KEYWORDS if kw.lower() in text.lower()]

    def check_article_for_pdf(self, article_url):
        try:
            self.driver.get(article_url)
            time.sleep(2)
            page_source = self.driver.page_source

            article_date = None
            date_text    = ""
            month_pat = (r'\d{1,2}\s+(?:January|February|March|April|May|June|July|'
                         r'August|September|October|November|December)\s+\d{4}')

            # 優先抓署名行的日期，例如「23 May 2014, Gateway House」
            # （頁首會顯示今天日期，不能直接取第一個）
            candidates = re.findall(rf'({month_pat}),\s*Gateway House', page_source)
            # 找不到署名行時，改抓標題 <h1> 之前、最接近標題的日期
            if not candidates:
                h1_pos = page_source.find("<h1")
                if h1_pos != -1:
                    candidates = list(reversed(re.findall(rf'({month_pat})', page_source[:h1_pos])))

            for date_str in candidates:
                temp_date = self.parse_article_date(date_str)
                if temp_date:
                    date_text    = date_str
                    article_date = temp_date
                    break

            if not article_date:
                return None, [], None
            if article_date < self.start_date or article_date > self.end_date:
                return date_text, [], None

            try:
                title = self.driver.find_element(By.TAG_NAME, "h1").text.strip()
            except:
                title = ""

            content = ""
            try:
                content = self.driver.find_element(By.CLASS_NAME, "entry-content").text
            except:
                try:
                    content = self.driver.find_element(By.TAG_NAME, "article").text
                except:
                    content = self.driver.find_element(By.TAG_NAME, "body").text

            found_keywords = self.find_keywords_in_text(title + " " + content)
            if not found_keywords:
                return date_text, [], None

            pdf_url = None
            try:
                for link in self.driver.find_elements(By.TAG_NAME, 'a'):
                    href = (link.get_attribute('href') or '').lower()
                    text = (link.text or '').lower()
                    if href.endswith('.pdf') and 'gatewayhouse.in' in href:
                        pdf_url = link.get_attribute('href')
                        break
                    if href.endswith('.pdf') and any(kw in text for kw in ['download', 'pdf', 'full text']):
                        pdf_url = link.get_attribute('href')
                        break
            except:
                pass

            return date_text, found_keywords, pdf_url

        except Exception as e:
            return None, [], None

    def scrape_page(self, url):
        consecutive_old = 0
        page_num = 1

        print(f"\n正在爬取: {url}")
        print("-" * 60)

        while page_num <= 30:  # 翻頁上限，避免無限翻頁
            try:
                current_url = url if page_num == 1 else f"{url}page/{page_num}/"
                self.driver.get(current_url)
                time.sleep(3)
                print(f"  📄 第 {page_num} 頁...")

                articles = self.driver.find_elements(By.TAG_NAME, "article")
                if not articles:
                    break

                article_links = []
                for article in articles:
                    try:
                        for selector in ["h1.entry-title a", "h3 a", "h2 a"]:
                            elements = article.find_elements(By.CSS_SELECTOR, selector)
                            if elements:
                                href = elements[0].get_attribute("href")
                                if href and "gatewayhouse.in" in href:
                                    article_links.append(href)
                                break
                    except:
                        continue

                article_links = list(dict.fromkeys(article_links))  # 去重但保留原本順序
                if not article_links:
                    break

                print(f"  ✓ 找到 {len(article_links)} 篇文章\n")

                for idx, article_url in enumerate(article_links, 1):
                    print(f"    [{idx}/{len(article_links)}] 檢查: {article_url}")
                    date_text, keywords, pdf_url = self.check_article_for_pdf(article_url)

                    if not date_text:
                        print("      → ✗ 無法取得日期\n")
                        continue

                    article_date = self.parse_article_date(date_text)

                    if article_date and article_date < self.start_date:
                        consecutive_old += 1
                        print(f"      → ✗ 日期過早 [連續{consecutive_old}篇]\n")
                        if consecutive_old >= 10:
                            print(f"  ⚠ 連續 10 篇早於起始日期，停止此分類")
                            return
                        continue
                    else:
                        consecutive_old = 0

                    if keywords and pdf_url:
                        try:
                            title = self.driver.find_element(By.TAG_NAME, "h1").text.strip()
                        except:
                            title = article_url
                        formatted_date = article_date.strftime("%Y/%m/%d")
                        self.articles.append({
                            "機構": self.INSTITUTION,
                            "日期": formatted_date,
                            "標題": title,
                            "網址": article_url,
                            "關鍵字": ", ".join(keywords)
                        })
                        print(f"      → ✓✓✓ 已收錄！\n")
                    else:
                        print(f"      → ✗ 不符合\n")

                # 翻頁
                next_page_num = page_num + 1
                self.driver.get(f"{url}page/{next_page_num}/")
                time.sleep(2)
                test_articles = self.driver.find_elements(By.TAG_NAME, "article")
                if test_articles:
                    page_num = next_page_num
                else:
                    break

            except Exception as e:
                print(f"  ❌ 爬取錯誤: {str(e)}")
                break

    def run(self):
        print("=" * 60)
        print("Indian Council on Global Relations 爬蟲程式")
        print("=" * 60)
        try:
            for url in self.URLS:
                self.scrape_page(url)
        except Exception as e:
            print(f"\n❌ 程式執行錯誤: {str(e)}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.articles


# ════════════════════════════════════════════════════════════
#  JIIA Scraper
# ════════════════════════════════════════════════════════════

class JIIAScraper:
    KEYWORDS    = ["China", "Taiwan", "Taipei"]
    INSTITUTION = "Japan Institute of International Affairs"

    SECTIONS = [
        {"name": "Research Reports",  "base_url": "https://www.jiia.or.jp/eng/report/research-report/",   "page_url_fn": lambda p: f"https://www.jiia.or.jp/eng/report/research-report/index_{p}.html",   "type": "html"},
        {"name": "Strategic Comment", "base_url": "https://www.jiia.or.jp/eng/report/strategic_comment/", "page_url_fn": lambda p: f"https://www.jiia.or.jp/eng/report/strategic_comment/index_{p}.html", "type": "html"},
        {"name": "AJISS-Commentary",  "base_url": "https://www.jiia.or.jp/eng/report/ajiss_commentary/",  "page_url_fn": lambda p: f"https://www.jiia.or.jp/eng/report/ajiss_commentary/index_{p}.html",  "type": "html"},
        {"name": "Policy Brief",      "base_url": "https://www.jiia.or.jp/eng/report/policy_brief/",      "page_url_fn": lambda p: f"https://www.jiia.or.jp/eng/report/policy_brief/index_{p}.html",      "type": "pdf"},
        {"name": "International Affairs", "base_url": "https://www.jiia.or.jp/eng/report/international_affairs/", "page_url_fn": lambda p: f"https://www.jiia.or.jp/eng/report/international_affairs/index_{p}.html", "type": "pdf"},
    ]

    def __init__(self, start_date, end_date):
        self.driver     = make_driver()
        self.results    = []
        self.start_date = start_date
        self.end_date   = end_date

    @staticmethod
    def parse_date(text):
        if not text:
            return None
        m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', text)
        if m:
            try:
                return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
        return None

    def get_cards_from_listing(self, list_url):
        self.driver.get(list_url)
        time.sleep(2)
        raw_cards = self.driver.execute_script("""
            const cards = [];
            document.querySelectorAll('a').forEach(a => {
                const h2 = a.querySelector('h2');
                if (!h2) return;
                const href  = a.href || '';
                const title = h2.textContent.trim();
                const timeEl = a.querySelector('[class*="p-article-list-01__time"]');
                let dateText = '';
                if (timeEl) {
                    dateText = timeEl.getAttribute('data-original') || timeEl.textContent.trim();
                }
                if (!dateText || !/\\d{2}\\.\\d{2}\\.\\d{4}/.test(dateText)) {
                    const html = a.innerHTML || '';
                    const m = html.match(/\\d{2}\\.\\d{2}\\.\\d{4}/);
                    dateText = m ? m[0] : (dateText || '');
                }
                cards.push({ href, title, dateText });
            });
            return cards;
        """)
        cards = []
        for c in raw_cards:
            href = c.get("href", "")
            if not href:
                continue
            date_obj = self.parse_date(c.get("dateText", ""))
            cards.append((href, c.get("title", ""), date_obj))
        return cards

    def check_pdf_keywords(self, pdf_url):
        if not HAS_PDFPLUMBER:
            return None
        try:
            resp = requests.get(pdf_url, timeout=30)
            resp.raise_for_status()
            with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
                text = "".join(p.extract_text() or "" for p in pdf.pages)
            return [kw for kw in self.KEYWORDS if kw in text]
        except Exception:
            return None

    def process_article_html(self, url, card_date=None):
        self.driver.get(url)
        time.sleep(1.5)

        date_obj = card_date
        if date_obj is None:
            try:
                m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', self.driver.page_source)
                if m:
                    date_obj = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except:
                pass

        if date_obj is None:
            return None, "❌ 日期無法解析"
        if date_obj > self.end_date:
            return None, f"⏭ 晚於結束日期"
        if date_obj < self.start_date:
            return None, f"⏪ 早於起始日期"

        title = ""
        try:
            title = self.driver.find_element(By.CSS_SELECTOR, "h1").text.strip()
        except:
            pass

        has_pdf = bool(self.driver.find_elements(By.XPATH,
            "//ul[contains(@class,'p-btn-list')]//a[contains(@class,'c-btn-01') and contains(@aria-label,'PDF')]"))
        if not has_pdf:
            return None, "📄 無 PDF 下載按鈕"

        try:
            body_text = self.driver.find_element(By.TAG_NAME, "body").text
        except:
            body_text = title

        found_kw = [kw for kw in self.KEYWORDS if kw in body_text or kw in title]
        if not found_kw:
            return None, "🔍 無關鍵字"

        return {"institution": self.INSTITUTION, "date": date_obj, "title": title, "url": url, "keywords": ", ".join(found_kw)}, None

    def process_pdf_direct(self, url, title, card_date):
        if card_date is None:
            return None, "❌ 日期無法解析"
        found_kw = self.check_pdf_keywords(url)
        if found_kw is None:
            return None, "⚠ PDF 下載失敗"
        if not found_kw:
            return None, "🔍 無關鍵字"
        return {"institution": self.INSTITUTION, "date": card_date, "title": title, "url": url, "keywords": ", ".join(found_kw)}, None

    def run_pagination(self, section):
        name, base_url, page_url_fn, sec_type = section["name"], section["base_url"], section["page_url_fn"], section["type"]
        print(f"\n{'-' * 20}\n{name}\n{'-' * 20}")
        page, consecutive_old = 1, 0

        while True:
            list_url = base_url if page == 1 else page_url_fn(page)
            print(f"  📄 列表頁 {page}")
            cards = self.get_cards_from_listing(list_url)
            if not cards:
                break

            stop = False
            for url, title, card_date in cards:
                if card_date is not None:
                    if card_date > self.end_date:
                        continue
                    if card_date < self.start_date:
                        consecutive_old += 1
                        if consecutive_old >= 3:
                            stop = True
                            break
                        continue

                if sec_type == "html":
                    info, reason = self.process_article_html(url, card_date)
                else:
                    info, reason = self.process_pdf_direct(url, title, card_date)

                if info is None:
                    if reason and "早於" in reason:
                        consecutive_old += 1
                        if consecutive_old >= 3:
                            stop = True
                            break
                    continue

                consecutive_old = 0
                self.results.append(info)
                print(f"  ✅ 收錄：{info['title'][:60]}")

            if stop:
                break
            page += 1
            time.sleep(1)

    def run(self):
        print("=" * 60)
        print("Japan Institute of International Affairs 文章爬蟲")
        print("=" * 60)
        try:
            for section in self.SECTIONS:
                self.run_pagination(section)
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()

        return [{"機構": r["institution"], "日期": r["date"].strftime("%Y/%m/%d"),
                 "標題": r["title"], "網址": r["url"], "關鍵字": r["keywords"]} for r in self.results]


# ════════════════════════════════════════════════════════════
#  Lowy Institute Scraper
# ════════════════════════════════════════════════════════════

class LowyScraper:
    INSTITUTION             = "Lowy Institute for International Policy"
    KEYWORDS                = ["China", "Taiwan", "Taipei"]
    BASE_URL                = "https://www.lowyinstitute.org/publications"
    CONSECUTIVE_EARLY_LIMIT = 3

    DATE_FORMATS = [
        "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S",
        "%d %B %Y", "%B %d, %Y", "%d %b %Y", "%b %d, %Y",
        "%B %Y", "%b %Y", "%Y-%m-%d",
    ]

    def __init__(self, start_date, end_date):
        self.driver     = make_driver()
        self.start_date = start_date
        self.end_date   = end_date
        self.results    = []

    def try_parse(self, raw):
        raw = raw.strip()
        for fmt in self.DATE_FORMATS:
            try:
                dt = datetime.strptime(raw, fmt)
                return dt.replace(tzinfo=None) if dt.tzinfo else dt
            except ValueError:
                pass
        for fmt in self.DATE_FORMATS:
            try:
                dt = datetime.strptime(raw[:19], fmt)
                return dt.replace(tzinfo=None) if dt.tzinfo else dt
            except ValueError:
                pass
        return None

    def parse_listing_page(self):
        articles  = []
        seen_urls = set()

        try:
            links = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='/publications/']")
            for lnk in links:
                href = lnk.get_attribute("href") or ""

                if not href or href.rstrip("/") == self.BASE_URL:
                    continue
                if "?" in href or "#" in href:
                    continue
                if href in seen_urls:
                    continue

                # 從連結內部找 h2/h3
                title = ""
                try:
                    h     = lnk.find_element(By.CSS_SELECTOR, "h2, h3")
                    title = h.text.strip()
                except NoSuchElementException:
                    pass

                if not title:
                    title = lnk.text.strip()

                # 純圖片連結：往父層找 h2/h3
                if not title:
                    try:
                        parent = self.driver.execute_script(
                            "return arguments[0].closest('article, li, div');", lnk)
                        if parent:
                            h     = parent.find_element(By.CSS_SELECTOR, "h2, h3")
                            title = h.text.strip()
                    except Exception:
                        pass

                if not title or not href:
                    continue

                # Podcast 過濾
                art_type = ""
                try:
                    parent = self.driver.execute_script(
                        "return arguments[0].closest('article, li, div');", lnk)
                    if parent:
                        for tag_sel in ["span.tag", "div.card__tag",
                                        "[class*='type']", "[class*='tag']"]:
                            try:
                                art_type = parent.find_element(
                                    By.CSS_SELECTOR, tag_sel).text.strip().lower()
                                break
                            except NoSuchElementException:
                                pass
                except Exception:
                    pass

                if "podcast" in art_type:
                    continue

                seen_urls.add(href)
                articles.append({"title": title, "url": href})

        except Exception:
            pass

        return articles

    def get_article_info(self, url):
        try:
            self.driver.get(url)
            time.sleep(5)
        except Exception:
            return None

        # ── 日期 ──
        pub_date = None
        for sel in ["div.date", "span.date-display-single", "time", "p.date",
                    "span[class*='date']", "div[class*='date']",
                    "meta[name='article:published_time']",
                    "meta[property='article:published_time']"]:
            try:
                el       = self.driver.find_element(By.CSS_SELECTOR, sel)
                raw      = el.get_attribute("datetime") or el.get_attribute("content") or el.text
                pub_date = self.try_parse(raw or "")
                if pub_date:
                    break
            except NoSuchElementException:
                continue

        if pub_date is None:
            try:
                body_text = self.driver.find_element(By.TAG_NAME, "body").text
                for pat in [r"(\d{1,2}\s+\w+\s+\d{4})", r"(\w+\s+\d{1,2},\s+\d{4})"]:
                    m = re.search(pat, body_text)
                    if m:
                        pub_date = self.try_parse(m.group(1))
                        if pub_date:
                            break
            except Exception:
                pass

        if pub_date is None:
            return None

        # ── 標題 ──
        try:
            title = self.driver.find_element(By.CSS_SELECTOR, "h1").text.strip()
        except Exception:
            title = url

        # ── 關鍵字：只搜尋內容區，排除 header/nav/banner ──
        search_text = ""
        for sel in ["main", "article", "div.main-content",
                    "div[class*='content']", "div[class*='body']"]:
            try:
                el          = self.driver.find_element(By.CSS_SELECTOR, sel)
                search_text = el.text
                if search_text.strip():
                    break
            except NoSuchElementException:
                continue

        if not search_text.strip():
            search_text = title

        found_keywords = [kw for kw in self.KEYWORDS
                          if re.search(kw, search_text, re.IGNORECASE)]

        # ── Download 按鈕偵測（三層）──
        has_pdf = False

        # 方法1：<a> 文字含 "download" 或 href 為 .pdf
        if not has_pdf:
            try:
                for lnk in self.driver.find_elements(By.TAG_NAME, "a"):
                    link_text = lnk.text.strip().lower()
                    href_val  = (lnk.get_attribute("href") or "").lower()
                    if "download" in link_text or href_val.endswith(".pdf"):
                        has_pdf = True
                        break
            except Exception:
                pass

        # 方法2：<span> 文字完全等於 "Download"（Lowy sr-only span 結構）
        if not has_pdf:
            try:
                for sp in self.driver.find_elements(By.TAG_NAME, "span"):
                    if sp.text.strip().lower() == "download":
                        has_pdf = True
                        break
            except Exception:
                pass

        # 方法3：CSS selector 兜底
        if not has_pdf:
            for sel in ["a.btn-pdf", "a[href$='.pdf']",
                        "a[class*='pdf']", "a[class*='download']"]:
                try:
                    if self.driver.find_elements(By.CSS_SELECTOR, sel):
                        has_pdf = True
                        break
                except Exception:
                    pass

        return {"date": pub_date, "title": title, "url": url,
                "keywords": found_keywords, "has_pdf": has_pdf}

    def scrape_publications_main(self):
        print(f"\n{'='*20}")
        print("爬取 Publications 主頁（逐篇進入）")
        print(f"{'='*20}")

        page              = 0
        consecutive_early = 0
        visited_urls      = set()

        while True:
            url = self.BASE_URL if page == 0 else f"{self.BASE_URL}?page={page}"
            print(f"\n  第 {page+1} 頁：{url}")
            self.driver.get(url)
            time.sleep(6)

            articles = self.parse_listing_page()
            if not articles:
                print("  → 找不到更多文章，結束")
                break

            stop_all = False
            for art in articles:
                art_url = art["url"]
                if art_url in visited_urls:
                    continue

                print(f"  ▶ 進入：{art['title'][:60]}")
                info = self.get_article_info(art_url)
                visited_urls.add(art_url)

                if info is None:
                    print("    ✗ 無法取得日期，跳過")
                    continue

                pub_date = info["date"]
                date_str = pub_date.strftime("%Y/%m/%d")
                print(f"    日期：{date_str}  關鍵字：{info['keywords']}  PDF：{info['has_pdf']}")

                if pub_date < self.start_date:
                    consecutive_early += 1
                    print(f"    ✗ 早於起始時間 (連續第 {consecutive_early} 篇)")
                    if consecutive_early >= self.CONSECUTIVE_EARLY_LIMIT:
                        print(f"  → 連續 {self.CONSECUTIVE_EARLY_LIMIT} 篇早於起始時間，停止")
                        stop_all = True
                        break
                    continue
                else:
                    consecutive_early = 0

                if pub_date > self.end_date:
                    print("    ✗ 晚於結束時間，跳過")
                    continue

                if info["keywords"] and info["has_pdf"]:
                    self.results.append({
                        "機構":  self.INSTITUTION,
                        "日期":  date_str,
                        "標題":  info["title"],
                        "網址":  art_url,
                        "關鍵字": ", ".join(info["keywords"]),
                    })
                    print("    ✓ 收錄！")

            if stop_all:
                break

            page += 1
            time.sleep(2)

    def run(self):
        print("=" * 60)
        print("Lowy Institute 文章爬蟲")
        print("1. 爬取 /publications 主頁（逐篇進入）")
        print("2. 跳過 Podcast 類型")
        print("3. 關鍵字只搜尋內容區，排除 header/nav/banner")
        print("4. Download 按鈕以三層方式偵測（含 sr-only span 結構）")
        print("5. 連續 3 篇發布日期早於起始日期即停止")
        print("=" * 60)
        try:
            self.scrape_publications_main()
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.results
        
# ════════════════════════════════════════════════════════════
#  ECIPE Scraper
# ════════════════════════════════════════════════════════════

class ECIPEScraper:
    INSTITUTION = "European Centre for International Political Economy"
    KEYWORDS    = ["China", "Taiwan", "Taipei"]
    CATEGORIES  = {
        "Publications":  {"base_url": "https://ecipe.org/browse/?subj_keyword={keyword}&subj_subject=&subj_year=&subj_order=recent&subj_meta_post_type=&subj_category=ecipepublications", "pagination": True},
        "Insights":      {"base_url": "https://ecipe.org/browse/?subj_keyword={keyword}&subj_subject=&subj_year=&subj_order=recent&subj_meta_post_type=&subj_category=post",              "pagination": True},
        "Articles":      {"base_url": "https://ecipe.org/browse/?subj_keyword={keyword}&subj_subject=&subj_year=&subj_order=recent&subj_meta_post_type=article&subj_category=ecipemediapost", "pagination": False},
        "Books or Papers": {"base_url": "https://ecipe.org/browse/?subj_keyword={keyword}&subj_subject=&subj_year=&subj_order=recent&subj_meta_post_type=bookorpaper&subj_category=ecipemediapost", "pagination": False},
    }

    def __init__(self, start_date, end_date):
        self.driver       = make_driver()
        self.results      = []
        self.visited_urls = set()
        self.start_date   = start_date
        self.end_date     = end_date

    @staticmethod
    def parse_date(date_str):
        try:
            if "-" in date_str and len(date_str) >= 10:
                return datetime.strptime(date_str[:10], "%Y-%m-%d")
        except:
            pass
        for fmt in ["%B %Y", "%d %B %Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"]:
            try:
                return datetime.strptime(date_str.strip(), fmt)
            except:
                continue
        return None

    def check_keywords_in_text(self, text, found_keyword):
        found = []
        text_lower = text.lower()
        for kw in self.KEYWORDS:
            if kw.lower() in text_lower:
                found.append(kw)
        return found

    def check_pdf_content(self, pdf_url):
        if not HAS_PYPDF2:
            return []
        try:
            response   = requests.get(pdf_url, timeout=30)
            pdf_reader = PyPDF2.PdfReader(BytesIO(response.content))
            text       = "".join(p.extract_text() or "" for p in pdf_reader.pages)
            return [kw for kw in self.KEYWORDS if kw.lower() in text.lower()]
        except:
            return []

    def scrape_category(self, base_url, category_name, keyword, need_pagination):
        print(f"\n開始爬取: {category_name} - {keyword}")
        page = 1

        while page <= 50:
            url = base_url if page == 1 else base_url.replace("/browse/", f"/browse/page/{page}/")
            self.driver.get(url)
            time.sleep(3)

            articles = self.driver.find_elements(By.CSS_SELECTOR, "li.post-listing-item, .post-listing-item")
            if not articles:
                break

            consecutive_old = 0
            for idx, article in enumerate(articles):
                try:
                    title, article_url = "", ""
                    for sel in ["header.post-listing-heading a", "h2 a", "h3 a", "a[href*='/publications/']"]:
                        try:
                            elem = article.find_element(By.CSS_SELECTOR, sel)
                            title = (elem.get_attribute("title") or elem.text).strip().replace(" - PDF download", "")
                            article_url = elem.get_attribute("href")
                            if title and article_url:
                                break
                        except:
                            continue

                    if not article_url or article_url in self.visited_urls:
                        continue

                    if category_name in ["Articles", "Books or Papers"]:
                        if not ("ecipe.org" in article_url or "geoii.eu" in article_url):
                            continue

                    date_str = ""
                    try:
                        date_elem = article.find_element(By.CSS_SELECTOR, "time, .date")
                        datetime_attr = date_elem.get_attribute("datetime")
                        date_str = (datetime_attr.split("T")[0] if datetime_attr and "T" in datetime_attr
                                    else (datetime_attr or date_elem.text.strip()))
                    except:
                        pass

                    article_date = self.parse_date(date_str)

                    if article_date:
                        if article_date < self.start_date:
                            consecutive_old += 1
                            if consecutive_old >= 3:
                                return
                            continue
                        elif article_date > self.end_date:
                            consecutive_old = 0
                            continue
                        else:
                            consecutive_old = 0

                    self.visited_urls.add(article_url)

                    if category_name in ["Articles", "Books or Papers"] and article_url.endswith(".pdf"):
                        keywords_found = self.check_pdf_content(article_url)
                        if keywords_found:
                            formatted_date = date_str.replace("-", "/") if date_str and "-" in date_str else date_str
                            self.results.append({
                                "機構": self.INSTITUTION, "日期": formatted_date,
                                "標題": title, "網址": article_url, "關鍵字": ", ".join(keywords_found)
                            })
                        continue

                    current_url = self.driver.current_url
                    self.driver.get(article_url)
                    time.sleep(2)

                    page_text = self.driver.find_element(By.TAG_NAME, "body").text
                    keywords_found = self.check_keywords_in_text(page_text, keyword)

                    if keywords_found:
                        has_pdf = False
                        for sel in ["a.button.bold.pdf_download", "a.pdf_download"]:
                            if self.driver.find_elements(By.CSS_SELECTOR, sel):
                                has_pdf = True
                                break
                        if not has_pdf:
                            for btn in self.driver.find_elements(By.XPATH, '//a[contains(translate(., "DOWNLOADPDF", "downloadpdf"), "downloadpdf")]'):
                                href = btn.get_attribute("href")
                                if href and ".pdf" in href.lower():
                                    has_pdf = True
                                    break

                        if has_pdf:
                            formatted_date = date_str.replace("-", "/") if date_str and "-" in date_str else date_str
                            self.results.append({
                                "機構": self.INSTITUTION, "日期": formatted_date,
                                "標題": title, "網址": article_url, "關鍵字": ", ".join(keywords_found)
                            })
                            print(f"  ✅ 收錄: {title[:50]}")

                    self.driver.get(current_url)
                    time.sleep(2)

                except Exception:
                    continue

            if not need_pagination:
                break
            page += 1

    def run(self):
        print("=" * 60)
        print("European Centre for International Political Economy文章爬蟲")
        print("=" * 60)
        try:
            for keyword in ["china", "taiwan", "taipei"]:
                for category_name, config in self.CATEGORIES.items():
                    url = config["base_url"].format(keyword=keyword)
                    self.scrape_category(url, category_name, keyword.capitalize(), config["pagination"])
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()

        dedup = {}
        for r in self.results:
            key = r["網址"]
            if key not in dedup:
                dedup[key] = r
        return list(dedup.values())


# ════════════════════════════════════════════════════════════
#  DGAP Scraper
# ════════════════════════════════════════════════════════════

class DGAPScraper:
    INSTITUTION      = "German Council on Foreign Relations"
    PUBLICATIONS_URL = "https://dgap.org/en/publications"
    KEYWORDS         = ["china", "taiwan", "taipei"]
    SKIP_TYPES       = {"IPQ", "IP", "IP SPECIAL"}

    def __init__(self, start_date, end_date):
        self.driver           = make_driver()
        self.start_date       = start_date
        self.end_date         = end_date
        self.results          = []
        self.visited_urls     = set()
        self.visited_card_keys = set()

    def scrape(self):
        wait = WebDriverWait(self.driver, 15)
        self.driver.get(self.PUBLICATIONS_URL)
        time.sleep(2)
        early_count = 0
        goto_export = False

        while True:
            cards = self.driver.find_elements(By.CSS_SELECTOR,
                "div.column.is-one-third-desktop.is-half-tablet.is-full")

            for card in cards:
                try:
                    time_el     = card.find_element(By.TAG_NAME, "time")
                    date_str    = time_el.get_attribute("datetime")
                    article_dt  = datetime.strptime(date_str[:10], "%Y-%m-%d")
                except:
                    early_count = 0
                    continue

                if article_dt > self.end_date:
                    early_count = 0
                    continue
                if article_dt < self.start_date:
                    early_count += 1
                    if early_count >= 3:
                        goto_export = True
                        break
                    continue
                else:
                    early_count = 0

                try:
                    _raw_title = card.find_element(By.CSS_SELECTOR, "h3").text.strip()
                except:
                    _raw_title = ""
                card_key = f"{date_str[:10]}|{_raw_title}"
                if card_key in self.visited_card_keys:
                    continue
                self.visited_card_keys.add(card_key)

                pub_type = ""
                try:
                    pub_type = card.find_element(By.CSS_SELECTOR, "span.publication-type__text").text.strip()
                except:
                    pass
                if pub_type.upper() in self.SKIP_TYPES:
                    continue

                try:
                    link_el     = card.find_element(By.CSS_SELECTOR, "a")
                    article_url = link_el.get_attribute("href")
                except:
                    continue

                if not article_url or article_url in self.visited_urls:
                    continue
                self.visited_urls.add(article_url)

                self.driver.execute_script("window.open(arguments[0]);", article_url)
                self.driver.switch_to.window(self.driver.window_handles[-1])
                time.sleep(1.5)

                found_keywords = []
                has_pdf        = False
                title_text     = ""

                try:
                    try:
                        title_text = self.driver.find_element(By.CSS_SELECTOR, "h1.field--name-field-title span, h1 span").text
                    except:
                        try:
                            title_text = self.driver.find_element(By.TAG_NAME, "h1").text
                        except:
                            title_text = ""

                    body_text = self.driver.find_element(By.TAG_NAME, "body").text
                    combined = (title_text + " " + body_text).lower()
                    for kw in self.KEYWORDS:
                        if kw in combined:
                            found_keywords.append(kw.capitalize())

                    pdf_links = self.driver.find_elements(By.CSS_SELECTOR,
                        "span.file.file--mime-application-pdf a, span.file--mime-application-pdf a")
                    if pdf_links:
                        has_pdf = True

                except Exception as e:
                    print(f"    ✗ 解析失敗：{e}")

                self.driver.close()
                self.driver.switch_to.window(self.driver.window_handles[0])
                time.sleep(0.5)

                if found_keywords and has_pdf:
                    self.results.append({
                        "機構": self.INSTITUTION, "日期": article_dt.strftime("%Y/%m/%d"),
                        "標題": title_text.strip(), "網址": article_url,
                        "關鍵字": ", ".join(found_keywords)
                    })
                    print(f"    ✓ 收錄！關鍵字：{', '.join(found_keywords)}")

            if goto_export:
                break

            try:
                prev_count = len(cards)
                load_more  = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                    "a.button__wrapper.button, a[data-drupal-views-infinite-scroll-pager], a.ajax-pager__item--next")))
                self.driver.execute_script("arguments[0].click();", load_more)
                WebDriverWait(self.driver, 15).until(lambda d: len(
                    d.find_elements(By.CSS_SELECTOR, "div.column.is-one-third-desktop.is-half-tablet.is-full")
                ) > prev_count)
                time.sleep(1)
            except:
                break

    def run(self):
        print("=" * 60)
        print("German Council on Foreign Relations 文章爬蟲")
        print("=" * 60)
        try:
            self.scrape()
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.results


# ════════════════════════════════════════════════════════════
#  Elcano Scraper
# ════════════════════════════════════════════════════════════

class ElcanoScraper:
    INSTITUTION = "Elcano Royal Institute"
    KEYWORDS    = ["China", "Taiwan", "Taipei"]
    CATEGORIES  = [
        ("Work Document", "https://www.realinstitutoelcano.org/en/work-document/"),
        ("Analysis",      "https://www.realinstitutoelcano.org/en/analyses/"),
        ("Report",        "https://www.realinstitutoelcano.org/en/report/"),
        ("Policy Paper",  "https://www.realinstitutoelcano.org/en/policy-paper/"),
        ("Monograph",     "https://www.realinstitutoelcano.org/en/monographs/"),
        ("Survey",        "https://www.realinstitutoelcano.org/en/surveys/"),
    ]
    PAGE_PATTERNS = {
        "https://www.realinstitutoelcano.org/en/work-document/": "https://www.realinstitutoelcano.org/en/work-document/page/%d/?pt=work_document&c=&y=",
        "https://www.realinstitutoelcano.org/en/analyses/":      "https://www.realinstitutoelcano.org/en/analyses/page/%d/?pt=analyses&c=&y=",
        "https://www.realinstitutoelcano.org/en/report/":        "https://www.realinstitutoelcano.org/en/report/page/%d/?pt=report&c=&y=",
        "https://www.realinstitutoelcano.org/en/policy-paper/":  "https://www.realinstitutoelcano.org/en/policy-paper/page/%d/?pt=policy-paper&c=&y=",
        "https://www.realinstitutoelcano.org/en/monographs/":    "https://www.realinstitutoelcano.org/en/monographs/page/%d/?pt=monographs&c=&y=",
        "https://www.realinstitutoelcano.org/en/surveys/":       "https://www.realinstitutoelcano.org/en/surveys/page/%d/?pt=surveys&c=&y=",
    }

    def __init__(self, start_date, end_date):
        self.start_date = start_date
        self.end_date   = end_date
        self.driver     = make_driver()
        self.wait       = WebDriverWait(self.driver, 15)
        self.results    = []

    @staticmethod
    def parse_date_str(s):
        s = s.strip()
        for fmt in ("%d %b %Y", "%d %B %Y", "%B %d, %Y", "%d/%m/%Y"):
            try:
                return datetime.strptime(s, fmt)
            except:
                continue
        m = re.search(r"(\d{1,2}\s+\w+\s+\d{4})", s)
        if m:
            try:
                return datetime.strptime(m.group(1), "%d %b %Y")
            except:
                pass
        return None

    def find_keywords(self, text):
        return [kw for kw in self.KEYWORDS if re.search(kw, text, re.IGNORECASE)]

    def scrape_article(self, url):
        self.driver.get(url)
        time.sleep(1.5)
        title_text = ""
        try:
            title_text = self.driver.find_element(By.CSS_SELECTOR, "h1.entry-title").text
        except:
            pass
        body_text = ""
        for sel in ["div.entry-content.maincol", "div.entry-content", "div.maincol"]:
            try:
                body_text = self.driver.find_element(By.CSS_SELECTOR, sel).text
                break
            except:
                continue
        found_kw = self.find_keywords(title_text + " " + body_text)
        has_pdf = False
        try:
            dl_section = self.driver.find_element(By.CSS_SELECTOR, "section.widget.downloads")
            for lnk in dl_section.find_elements(By.TAG_NAME, "a"):
                href = lnk.get_attribute("href") or ""
                if ".pdf" in href.lower() or "pdf" in lnk.text.lower():
                    has_pdf = True
                    break
        except:
            pass
        return found_kw, has_pdf

    def scrape(self):
        for cat_name, base_url in self.CATEGORIES:
            print(f"\n{'-'*20}\n▶ 分類：{cat_name}\n{'-'*20}")
            page_num = 1
            stop_category = False

            while not stop_category:
                page_url = base_url if page_num == 1 else self.PAGE_PATTERNS[base_url] % page_num
                self.driver.get(page_url)
                time.sleep(2)

                try:
                    self.wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "article.elcano-card, article[class*='elcano-card']")))
                except:
                    break

                cards = self.driver.find_elements(By.CSS_SELECTOR, "article.elcano-card, article[class*='elcano-card']")
                if not cards:
                    break

                consecutive_old = 0
                for card in cards:
                    article_dt = None
                    try:
                        date_el = card.find_element(By.CSS_SELECTOR, "span.posted-on")
                        article_dt = self.parse_date_str(date_el.get_attribute("title") or date_el.text)
                    except:
                        pass

                    if article_dt is not None:
                        if article_dt < self.start_date:
                            consecutive_old += 1
                            if consecutive_old >= 3:
                                stop_category = True
                                break
                            continue
                        else:
                            consecutive_old = 0
                        if article_dt > self.end_date:
                            continue

                    try:
                        article_url = card.find_element(By.CSS_SELECTOR, "h2.entry-title a").get_attribute("href")
                    except:
                        continue
                    if not article_url:
                        continue

                    card_title = ""
                    try:
                        card_title = card.find_element(By.CSS_SELECTOR, "h2.entry-title").text.strip()
                    except:
                        pass

                    found_kw, has_pdf = self.scrape_article(article_url)

                    if found_kw and has_pdf:
                        try:
                            final_title = self.driver.find_element(By.CSS_SELECTOR, "h1.entry-title").text.strip()
                        except:
                            final_title = card_title

                        self.results.append({
                            "機構": self.INSTITUTION,
                            "日期": article_dt.strftime("%Y/%m/%d") if article_dt else "",
                            "標題": final_title,
                            "網址": article_url,
                            "關鍵字": ", ".join(sorted(set(found_kw)))
                        })
                        print(f"  ✅ 收錄：{final_title[:60]}")

                    self.driver.back()
                    time.sleep(1.5)

                if not stop_category:
                    page_num += 1

    def run(self):
        print("=" * 60)
        print("Elcano Royal Institute 文章爬蟲")
        print("=" * 60)
        try:
            self.scrape()
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.results


# ════════════════════════════════════════════════════════════
#  KAS Scraper
# ════════════════════════════════════════════════════════════

class KASScraper:
    INSTITUTION = "Konrad-Adenauer-Stiftung"
    BASE_URL    = "https://www.kas.de/de/publikationen"
    KEYWORDS    = ["China", "Taiwan", "Taipei"]

    def __init__(self, start_date, end_date):
        self.start_date              = start_date
        self.end_date                = end_date
        self.articles_data           = []
        self.driver                  = make_driver()
        self.wait                    = WebDriverWait(self.driver, 10)
        self.consecutive_old_articles = 0

    @staticmethod
    def parse_german_date(date_text):
        month_mapping = {
            "Januar": 1, "Februar": 2, "März": 3, "April": 4,
            "Mai": 5, "Juni": 6, "Juli": 7, "August": 8,
            "September": 9, "Oktober": 10, "November": 11, "Dezember": 12,
        }
        try:
            date_text = date_text.strip()
            match = re.match(r"(\d+)\.\s+(\w+)\s+(\d{4})", date_text)
            if match:
                day, month_name, year = int(match.group(1)), match.group(2), int(match.group(3))
                if month_name in month_mapping:
                    return datetime(year, month_mapping[month_name], day)
        except:
            pass
        return None

    def check_keywords(self, text):
        return [kw for kw in self.KEYWORDS if kw.lower() in text.lower()]

    def has_pdf_download(self):
        try:
            buttons = self.driver.find_elements(By.XPATH,
                "//a[contains(text(), 'Herunterladen') or contains(@class, 'download') or contains(@href, '.pdf')]")
            return bool(buttons)
        except:
            return False

    def extract_article_info(self, article_url):
        try:
            self.driver.get(article_url)
            time.sleep(2)

            article_date = None
            for selector in ["//span[contains(@class, 'o-metadata--date')]", "//time"]:
                try:
                    for elem in self.driver.find_elements(By.XPATH, selector):
                        article_date = self.parse_german_date(elem.text.strip())
                        if article_date:
                            break
                    if article_date:
                        break
                except:
                    continue

            if not article_date:
                return None

            if article_date < self.start_date:
                self.consecutive_old_articles += 1
                return None
            elif article_date > self.end_date:
                self.consecutive_old_articles = 0
                return None
            else:
                self.consecutive_old_articles = 0

            try:
                title = self.driver.find_element(By.TAG_NAME, "h1").text.strip()
            except:
                title = "無標題"

            try:
                page_text = self.driver.find_element(By.TAG_NAME, "body").text
            except:
                page_text = ""

            keywords_found = self.check_keywords(title + " " + page_text)
            if not keywords_found:
                return None

            if not self.has_pdf_download():
                return None

            return {"機構": self.INSTITUTION, "日期": article_date.strftime("%Y/%m/%d"),
                    "標題": title, "網址": article_url, "關鍵字": ", ".join(keywords_found)}
        except:
            return None

    def get_article_links(self):
        try:
            time.sleep(3)
            article_links = []
            for elem in self.driver.find_elements(By.XPATH, "//a[contains(@href, '/detail/')]"):
                href = elem.get_attribute("href")
                if href and href.startswith("http"):
                    article_links.append(href)
            seen = set()
            return [l for l in article_links if not (l in seen or seen.add(l))]
        except:
            return []

    def click_next_page(self, current_page):
        try:
            next_page = current_page + 1
            next_url  = (f"{self.BASE_URL}?p_p_id=com_liferay_asset_publisher_web_portlet_"
                         f"AssetPublisherPortlet_INSTANCE_PUBLIKATIONEN&p_p_lifecycle=0"
                         f"&p_p_state=normal&p_p_mode=view&p_r_p_resetCur=false"
                         f"&_com_liferay_asset_publisher_web_portlet_AssetPublisherPortlet"
                         f"_INSTANCE_PUBLIKATIONEN_delta=10"
                         f"&_com_liferay_asset_publisher_web_portlet_AssetPublisherPortlet"
                         f"_INSTANCE_PUBLIKATIONEN_cur={next_page}&selectedTab=0")
            self.driver.get(next_url)
            time.sleep(3)
            return True
        except:
            return False

    def run(self):
        print("=" * 60)
        print("KAS (Konrad-Adenauer-Stiftung) 網站爬蟲")
        print("=" * 60)
        try:
            self.driver.get(self.BASE_URL)
            time.sleep(3)
            current_page = 1

            while True:
                print(f"\n📄 正在處理第 {current_page} 頁")
                article_links = self.get_article_links()
                if not article_links:
                    break

                for i, article_url in enumerate(article_links, 1):
                    if self.consecutive_old_articles >= 7:
                        return self.articles_data

                    article_info = self.extract_article_info(article_url)
                    if article_info:
                        self.articles_data.append(article_info)
                        print(f"  ✅ 收錄 {article_info['標題'][:50]}")

                    self.driver.back()
                    time.sleep(2)

                if self.consecutive_old_articles >= 7:
                    break

                if not self.click_next_page(current_page):
                    break
                current_page += 1

        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()

        return self.articles_data


# ════════════════════════════════════════════════════════════
#  SWP Scraper
# ════════════════════════════════════════════════════════════

class SWPScraper:
    INSTITUTION = "Stiftung Wissenschaft und Politik"
    KEYWORDS    = ["China", "Taiwan", "Taipei"]
    CATEGORIES  = {
        "SWP Research Paper": "https://www.swp-berlin.org/en/publications?publication%5Bfilter%5D%5B%5D=swpPublication%3A%2F2231%2F3104%2F",
        "SWP Study":          "https://www.swp-berlin.org/en/publications?publication%5Bfilter%5D%5B%5D=swpPublication%3A%2F2231%2F3004%2F",
        "SWP Comment":        "https://www.swp-berlin.org/en/publications?publication%5Bfilter%5D%5B%5D=swpPublication%3A%2F2231%2F3103%2F",
    }

    def __init__(self, start_date, end_date):
        self.start_date      = start_date
        self.end_date        = end_date
        self.driver          = make_driver()
        self.results         = []
        self.processed_urls  = set()

    def check_keywords(self, text):
        return [kw for kw in self.KEYWORDS if kw.lower() in (text or "").lower()]

    def scrape_category(self, category_name, category_url):
        print(f"\n{'-'*20}\n開始處理分類: {category_name}\n{'-'*20}")
        self.driver.get(category_url)
        time.sleep(3)
        consecutive_old = 0

        while True:
            try:
                articles = self.driver.find_elements(By.CSS_SELECTOR, "li.search__result-item")
                if not articles:
                    break

                for idx, article in enumerate(articles):
                    try:
                        title_elem  = article.find_element(By.CSS_SELECTOR, "h2.search__result-item-title a")
                        title       = title_elem.text.strip()
                        article_url = title_elem.get_attribute("href")

                        if article_url in self.processed_urls:
                            continue

                        self.driver.execute_script("window.open('');")
                        self.driver.switch_to.window(self.driver.window_handles[1])
                        self.driver.get(article_url)
                        time.sleep(2)

                        article_date = None
                        try:
                            date_text  = self.driver.find_element(By.CSS_SELECTOR, "span.small-text").text.strip()
                            date_match = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', date_text)
                            if date_match:
                                day, month, year = date_match.groups()
                                article_date = datetime(int(year), int(month), int(day))
                        except:
                            pass

                        if not article_date:
                            try:
                                body_text = self.driver.find_element(By.TAG_NAME, "body").text
                                for day, month, year in re.findall(r'(\d{2})\.(\d{2})\.(\d{4})', body_text):
                                    if 2020 <= int(year) <= 2030 and 1 <= int(month) <= 12:
                                        article_date = datetime(int(year), int(month), int(day))
                                        break
                            except:
                                pass

                        self.driver.close()
                        self.driver.switch_to.window(self.driver.window_handles[0])

                        if not article_date:
                            self.processed_urls.add(article_url)
                            continue

                        if article_date < self.start_date:
                            consecutive_old += 1
                            if consecutive_old >= 3:
                                self.processed_urls.add(article_url)
                                return
                            self.processed_urls.add(article_url)
                            continue
                        elif article_date > self.end_date:
                            self.processed_urls.add(article_url)
                            continue

                        consecutive_old = 0

                        self.driver.execute_script("window.open('');")
                        self.driver.switch_to.window(self.driver.window_handles[1])
                        self.driver.get(article_url)
                        time.sleep(2)

                        page_text = title + " " + self.driver.find_element(By.TAG_NAME, "body").text
                        keywords_found = self.check_keywords(page_text)

                        has_pdf = bool(self.driver.find_elements(By.CSS_SELECTOR, "a[href*='.pdf']"))

                        self.driver.close()
                        self.driver.switch_to.window(self.driver.window_handles[0])

                        if keywords_found and has_pdf:
                            self.results.append({
                                "機構": self.INSTITUTION, "日期": article_date.strftime("%Y/%m/%d"),
                                "標題": title, "網址": article_url, "關鍵字": ", ".join(keywords_found)
                            })
                            print(f"  ✓ 收錄: {title[:50]}")

                        self.processed_urls.add(article_url)

                    except Exception as e:
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        continue

                try:
                    load_more_btn = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "button.button--big")))
                    self.driver.execute_script("arguments[0].scrollIntoView();", load_more_btn)
                    time.sleep(1)
                    load_more_btn.click()
                    time.sleep(3)
                except:
                    break

            except Exception as e:
                print(f"  ⚠ 錯誤: {e}")
                break

    def run(self):
        print("=" * 60)
        print("SWP 爬蟲程式")
        print("=" * 60)
        try:
            for category_name, category_url in self.CATEGORIES.items():
                self.scrape_category(category_name, category_url)
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.results


# ════════════════════════════════════════════════════════════
#  匯出 Excel & 主程式
# ════════════════════════════════════════════════════════════

def export_to_excel(articles, filename="combined_articles_G4_normal.xlsx"):
    if not articles:
        print("\n⚠ 沒有符合條件的文章可以匯出")
        return 0

    wb = Workbook()
    ws = wb.active
    ws.title = "文章列表"
    headers = ["機構", "日期", "標題", "網址", "關鍵字"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font      = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for article in articles:
        ws.append([article["機構"], article["日期"], article["標題"], article["網址"], article["關鍵字"]])
    wb.save(filename)
    print(f"\n✓ Excel 已儲存: {filename}（共 {len(articles)} 筆）")
    return len(articles)


def main():
    all_articles = []

    scrapers = [
        ("GatewayHouse",   lambda: GatewayHouseScraper(start_date, end_date).run()),
        ("JIIA",           lambda: JIIAScraper(start_date, end_date).run()),
        ("Lowy Institute", lambda: LowyScraper(start_date, end_date).run()),
        ("ECIPE",          lambda: ECIPEScraper(start_date, end_date).run()),
        ("DGAP",           lambda: DGAPScraper(start_date, end_date).run()),
        ("KAS",            lambda: KASScraper(start_date, end_date).run()),
        ("Elcano",         lambda: ElcanoScraper(start_date, end_date).run()),
        ("SWP",            lambda: SWPScraper(start_date, end_date).run()),
    ]

    for name, run_fn in scrapers:
        try:
            articles = run_fn()
            all_articles.extend(articles)
            print(f"  ✓ {name} 完成，本次收錄 {len(articles)} 篇")
        except Exception as e:
            print(f"  ❌ {name} 發生未預期錯誤，跳過：{e}")

    excel_path = "combined_articles_G4_normal.xlsx"
    total_count = export_to_excel(all_articles, excel_path)

    print("\n📧 正在寄送 Email...")
    log_text = logger.get_log()
    send_email(excel_path, log_text, total_count)

    print("🎉 全部完成！")


if __name__ == "__main__":
    main()
