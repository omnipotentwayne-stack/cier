#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合併爬蟲：IFRI + ECFR + IDS + IISS + EVC + Sinopsis + ERIA + JETRO
自動計算日期區間（1/11/21號），headless模式，執行完寄信
"""

import io
import sys
import os
import re
import time
import calendar
import logging
import warnings
import traceback
import random
from datetime import datetime, date, timedelta
from typing import List, Dict, Tuple

import requests
import pandas as pd
from bs4 import BeautifulSoup
from dateutil import parser as dateparser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    StaleElementReferenceException
)
from openpyxl import Workbook
from openpyxl.styles import Font

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service as ChromeService
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False

logging.getLogger("pdfminer").setLevel(logging.ERROR)
warnings.filterwarnings("ignore")


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
    msg["Subject"] = f"G5爬蟲結果 {today.strftime('%Y/%m/%d')}（區間 {start_date.date()} ~ {end_date.date()}）"

    body = (
        f"G5爬蟲執行完畢。\n"
        f"執行日期：{today.strftime('%Y/%m/%d %H:%M')}\n"
        f"資料區間：{start_date.date()} ~ {end_date.date()}\n"
        f"共收錄：{total_count} 筆\n\n"
        f"詳細 log 請見附件 G5_log.txt，Excel 結果請見附件。"
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
    log_part.add_header("Content-Disposition", "attachment; filename=G5_log.txt")
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
        service = ChromeService(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=opts)
    return webdriver.Chrome(options=opts)


# ============================================================
# 共用輸出 Excel
# ============================================================

def save_to_excel(all_articles, output_file='combined_articles_G5_normal.xlsx'):
    if not all_articles:
        print("\n⚠️  沒有找到符合條件的文章")
        return 0
    df = pd.DataFrame(all_articles)
    df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"✓ 數據已保存到: {output_file}（共 {len(all_articles)} 筆）")
    return len(all_articles)


# ============================================================
# IFRI 爬蟲
# ============================================================

class IFRIScraper:
    def __init__(self):
        self.driver = None
        self.base_url = "https://www.ifri.org/en/research/search-publications"
        self.url_params = "field_date_de_publication%5Bmin%5D=1936&field_date_de_publication%5Bmax%5D=2026&field_type_publication%5B9007%5D=9007&field_type_publication%5B4%5D=4&field_type_publication%5B21%5D=21&field_type_publication%5B2356%5D=2356&field_type_publication%5B2412%5D=2412&field_type_publication%5B4425%5D=4425"
        self.articles_data = []
        self.keywords = ['China', 'Taiwan', 'Taipei']

    def parse_article_date(self, date_text):
        try:
            return datetime.strptime(' '.join(date_text.split()), '%d %B %Y')
        except:
            return None

    def check_keywords(self, text):
        return [kw for kw in self.keywords if kw.upper() in text.upper()]

    def scrape_article_details(self, article_url, article_date):
        try:
            self.driver.get(article_url)
            time.sleep(2)
            title = None
            for sel in ['h2.t-color-blue4', 'h2', 'h1']:
                try:
                    title = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                    if title:
                        break
                except:
                    continue
            if not title:
                return None

            page_text = ""
            for sel in ['.paragraph-analyse-content', '.field--name-body', '.node__content']:
                try:
                    page_text = self.driver.find_element(By.CSS_SELECTOR, sel).text
                    if len(page_text) > 100:
                        break
                except:
                    continue
            if not page_text:
                try:
                    page_text = self.driver.find_element(By.TAG_NAME, 'body').text
                except:
                    pass

            found_keywords = self.check_keywords(title + " " + page_text)
            if not found_keywords:
                return None

            try:
                pdf_buttons = self.driver.find_elements(By.XPATH,
                    "//a[contains(translate(., 'DOWNLOAD', 'download'), 'download') and contains(., '.pdf')]")
                if not pdf_buttons:
                    return None
            except:
                return None

            return {
                '機構': 'French Institute of International Relations',
                '日期': article_date.strftime('%Y/%m/%d'),
                '標題': title,
                '網址': article_url,
                '關鍵字': ', '.join(found_keywords)
            }
        except Exception as e:
            return None

    def scrape_page(self, page_num, start_dt, end_dt):
        url = f"{self.base_url}?{self.url_params}&page={page_num}"
        print(f"\n正在爬取第 {page_num + 1} 頁...")
        try:
            self.driver.get(url)
            time.sleep(3)
            articles = self.driver.find_elements(By.CSS_SELECTOR, 'div.views-row')
            if not articles:
                return True, 0

            article_info_list = []
            for article in articles:
                date_text = None
                try:
                    date_containers = article.find_elements(By.CSS_SELECTOR,
                        'div[class*="u-flex"][class*="u-align-items-center"][class*="u-mb10"]')
                    for dc in date_containers:
                        for div in dc.find_elements(By.XPATH, './div'):
                            for line in div.text.split('\n'):
                                line = line.strip()
                                if line and any(m in line for m in [
                                    'January','February','March','April','May','June',
                                    'July','August','September','October','November','December']):
                                    date_text = line
                                    break
                        if date_text:
                            break
                except:
                    continue
                if not date_text:
                    continue
                article_date = self.parse_article_date(date_text)
                if not article_date:
                    continue
                try:
                    link = article.find_element(By.CSS_SELECTOR, 'h4 a')
                    article_url = link.get_attribute('href')
                    if not article_url.startswith('http'):
                        article_url = 'https://www.ifri.org' + article_url
                except:
                    continue
                article_info_list.append({'url': article_url, 'date': article_date})

            early_count = 0
            for info in article_info_list:
                article_date = info['date']
                article_url = info['url']
                if article_date < start_dt:
                    early_count += 1
                    if early_count >= 3:
                        return False, early_count
                    continue
                if article_date > end_dt:
                    early_count = 0
                    continue
                early_count = 0
                article_data = self.scrape_article_details(article_url, article_date)
                if article_data:
                    self.articles_data.append(article_data)
                    print(f"  ✓ 收錄: {article_data['標題'][:50]}")
                self.driver.get(url)
                time.sleep(2)

            return True, early_count
        except Exception as e:
            print(f"❌ 頁面錯誤: {e}")
            return True, 0

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("French Institute of International Relations 網站爬蟲")
        print("=" * 60)
        self.driver = make_driver()
        try:
            page_num = 0
            while True:
                should_continue, _ = self.scrape_page(page_num, start_dt, end_dt)
                if not should_continue:
                    break
                page_num += 1
                time.sleep(1)
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.articles_data


# ============================================================
# ECFR 爬蟲
# ============================================================

class ECFRScraper:
    def __init__(self):
        self.driver = None
        self.articles_data = []
        self.visited_urls = set()

    def extract_date_from_article(self):
        try:
            for selector in [".post-meta time", "time[datetime]", "time", ".post-date"]:
                try:
                    date_elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for date_element in date_elements:
                        date_text = date_element.get_attribute('datetime') or date_element.text
                        if not date_text:
                            continue
                        if 'T' in date_text:
                            try:
                                return datetime.fromisoformat(date_text.split('T')[0])
                            except:
                                pass
                        for fmt in ['%Y-%m-%d', '%d %B %Y', '%B %d, %Y']:
                            try:
                                return datetime.strptime(date_text.strip().split('\n')[0], fmt)
                            except:
                                continue
                except:
                    continue
            return None
        except:
            return None

    def check_keywords(self):
        keywords = ['China', 'Taiwan', 'Taipei']
        try:
            page_text = self.driver.find_element(By.TAG_NAME, 'body').text
            return [kw for kw in keywords if kw.lower() in page_text.lower()]
        except:
            return []

    def has_download_pdf_button(self):
        try:
            for link in self.driver.find_elements(By.TAG_NAME, 'a'):
                try:
                    if not link.is_displayed():
                        continue
                    link_text = link.text.strip()
                    href = link.get_attribute('href') or ''
                    if 'download' in link_text.lower() and 'pdf' in link_text.lower() and '.pdf' in href.lower():
                        return True
                except:
                    continue
            return False
        except:
            return False

    def scrape_search_results(self, search_term, start_dt, end_dt):
        base_url = f"https://ecfr.eu/?s={search_term}"
        page = 1
        consecutive_old = 0

        while True:
            url = base_url if page == 1 else f"{base_url}&page={page + 1}"
            print(f"\n🔍 頁面: {url}")
            self.driver.get(url)
            time.sleep(3)

            articles = []
            for selector in ["article.tease", ".card-main-link-container", "article", ".tease"]:
                articles = self.driver.find_elements(By.CSS_SELECTOR, selector)
                if articles:
                    break

            if not articles:
                break

            for idx, article in enumerate(articles, 1):
                try:
                    article_html = article.get_attribute('outerHTML').lower()
                    if 'podcast' in article_html or 'event' in article_html:
                        continue

                    article_url = None
                    for sel in ["h2 a, h3 a, .post-title a", "a[href*='/article/']", "a[href*='/publication/']"]:
                        try:
                            link_element = article.find_element(By.CSS_SELECTOR, sel)
                            article_url = link_element.get_attribute('href')
                            break
                        except:
                            continue

                    if not article_url or 'ecfr.eu' not in article_url:
                        continue
                    if article_url in self.visited_urls:
                        continue
                    self.visited_urls.add(article_url)

                    self.driver.execute_script("window.open(arguments[0], '_blank');", article_url)
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    time.sleep(2)

                    article_date = self.extract_date_from_article()
                    if article_date:
                        if article_date < start_dt:
                            consecutive_old += 1
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                            if consecutive_old >= 3:
                                return False
                            continue
                        if article_date > end_dt:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                            continue
                        consecutive_old = 0
                    else:
                        self.driver.close()
                        self.driver.switch_to.window(self.driver.window_handles[0])
                        continue

                    keywords_found = self.check_keywords()
                    if not keywords_found or not self.has_download_pdf_button():
                        self.driver.close()
                        self.driver.switch_to.window(self.driver.window_handles[0])
                        continue

                    title = ""
                    for sel in ["h1", ".post-title", ".entry-title"]:
                        try:
                            title = self.driver.find_element(By.CSS_SELECTOR, sel).text
                            if title:
                                break
                        except:
                            continue

                    self.articles_data.append({
                        '機構': 'European Council on Foreign Relations',
                        '日期': article_date.strftime('%Y/%m/%d'),
                        '標題': title,
                        '網址': article_url,
                        '關鍵字': ', '.join(keywords_found)
                    })
                    print(f"    ✨ 收錄: {title[:50]}")

                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])
                    time.sleep(1)

                except Exception as e:
                    try:
                        if len(self.driver.window_handles) > 1:
                            self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                    except:
                        pass
                    continue

            page += 1
            time.sleep(2)
        return True

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("European Council on Foreign Relations 文章爬蟲")
        print("=" * 60)
        self.driver = make_driver()
        try:
            for search_term in ['china', 'taiwan', 'taipei']:
                print(f"\n{'-' * 20}\n搜尋: {search_term.upper()}\n{'-' * 20}")
                self.scrape_search_results(search_term, start_dt, end_dt)
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.articles_data


# ============================================================
# IDS 爬蟲
# ============================================================

IDS_DATE_RE = re.compile(r'^\d{1,2}\s+\w+\s+\d{4}$')

class IDSScraper:
    def __init__(self):
        self.driver = None
        self.articles_data = []
        self.keywords = ["China", "Taiwan", "Taipei"]
        self.institution = "Institute of Development Studies"
        self.base_url = "https://www.ids.ac.uk/research/publications/?current-page={page}#listing"

    def parse_article_date(self, text):
        text = text.strip()
        if not IDS_DATE_RE.match(text):
            return None
        try:
            return datetime.strptime(text, "%d %B %Y")
        except:
            return None

    def wait_for_articles(self, timeout=25):
        end = time.time() + timeout
        while time.time() < end:
            try:
                listing = self.driver.find_element(By.ID, "listing")
                candidates = listing.find_elements(By.CSS_SELECTOR, "p.c-content-item__date, p.ts-caption")
                if [el for el in candidates if IDS_DATE_RE.match(el.text.strip())]:
                    return True
            except:
                pass
            time.sleep(0.5)
        return False

    def get_listing_articles(self):
        articles = []
        try:
            listing = self.driver.find_element(By.ID, "listing")
        except:
            return articles
        candidates = listing.find_elements(By.CSS_SELECTOR, "p.c-content-item__date, p.ts-caption")
        date_els = [el for el in candidates if IDS_DATE_RE.match(el.text.strip())]
        for date_el in date_els:
            try:
                date_str = date_el.text.strip()
                date_obj = self.parse_article_date(date_str)
                link_el = None
                for level in range(1, 8):
                    try:
                        anc = date_el.find_element(By.XPATH, f"./ancestor::*[{level}]")
                        links = anc.find_elements(By.CSS_SELECTOR, "a[href*='/publications/']")
                        if links:
                            link_el = links[0]
                            break
                    except:
                        continue
                if not link_el:
                    continue
                url = link_el.get_attribute("href")
                title = link_el.text.strip()
                if not title:
                    for tag in ["h3", "h2", "h4"]:
                        try:
                            anc = date_el.find_element(By.XPATH, "./ancestor::*[3]")
                            h = anc.find_element(By.TAG_NAME, tag)
                            title = h.text.strip()
                            if title:
                                break
                        except:
                            continue
                if url and title:
                    articles.append((title, url, date_str, date_obj))
            except:
                continue
        seen = set()
        unique = []
        for item in articles:
            if item[1] not in seen:
                seen.add(item[1])
                unique.append(item)
        return unique

    def check_article(self, url):
        self.driver.get(url)
        time.sleep(2)
        has_access = False
        try:
            sidebar = self.driver.find_element(By.CSS_SELECTOR, "aside.c-right-sidebar, .c-tertiary-meta")
            if "access this publication" in sidebar.text.lower():
                has_access = True
        except:
            pass
        if not has_access:
            try:
                for h in self.driver.find_elements(By.CSS_SELECTOR, "h2"):
                    if "access this publication" in h.text.lower():
                        has_access = True
                        break
            except:
                pass

        title_text = ""
        try:
            title_text = self.driver.find_element(By.CSS_SELECTOR, "h1.c-single-header__heading, h1").text
        except:
            pass
        body_text = ""
        for sel in ["p.standfirst", "div.o-content-from-editor"]:
            try:
                body_text += " " + self.driver.find_element(By.CSS_SELECTOR, sel).text
            except:
                pass
        found_keywords = [kw for kw in self.keywords
                          if re.search(r'\b' + kw + r'\b', title_text + " " + body_text, re.IGNORECASE)]
        return found_keywords, has_access

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("Institute of Development Studies 文章爬蟲")
        print("=" * 60)
        self.driver = make_driver()
        page = 1
        try:
            while True:
                listing_url = self.base_url.format(page=page)
                print(f"\n📄 第 {page} 頁...")
                self.driver.get(listing_url)
                self.wait_for_articles(timeout=25)
                time.sleep(1.5)
                articles = self.get_listing_articles()
                if not articles:
                    break
                print(f"  ✅ 本頁 {len(articles)} 篇文章")
                early_count = 0
                stop_paging = False
                for title, art_url, date_str, date_obj in articles:
                    date_display = date_obj.strftime("%Y/%m/%d") if date_obj else "?"
                    if date_obj > end_dt:
                        early_count = 0
                        continue
                    if date_obj < start_dt:
                        early_count += 1
                        if early_count >= 3:
                            stop_paging = True
                            break
                        continue
                    early_count = 0
                    found_kws, has_access = self.check_article(art_url)
                    if found_kws and has_access:
                        self.articles_data.append({
                            "機構": self.institution, "日期": date_display,
                            "標題": title, "網址": art_url, "關鍵字": ", ".join(found_kws)
                        })
                        print(f"    ✅ 收錄！關鍵字: {', '.join(found_kws)}")
                    self.driver.get(listing_url)
                    time.sleep(2)
                    self.wait_for_articles(timeout=20)
                if stop_paging:
                    break
                page += 1
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            self.driver.quit()
        return self.articles_data


# ============================================================
# IISS 爬蟲
# ============================================================

class IISSScraper:
    def __init__(self):
        self.driver = None
        self.visited_urls = set()
        self.articles_data = []
        self.consecutive_old_articles = 0

    def parse_date(self, date_str):
        date_str = date_str.strip()
        date_str = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', date_str)
        for fmt in ['%d %B %Y', '%B %d %Y', '%Y/%m/%d', '%Y-%m-%d']:
            try:
                return datetime.strptime(date_str, fmt)
            except:
                continue
        return None

    def click_newest_sort(self):
        try:
            time.sleep(2)
            for by, selector in [
                (By.CSS_SELECTOR, "select.label--link"),
                (By.XPATH, "//select[.//option[@value='Newest']]"),
            ]:
                try:
                    select_element = self.driver.find_element(by, selector)
                    if select_element:
                        Select(select_element).select_by_value("Newest")
                        time.sleep(2)
                        return
                except:
                    continue
        except:
            pass

    def check_article_type(self, article_element):
        try:
            article_text = article_element.text.upper()
            for keyword in ['EVENT', 'TRANSCRIPT']:
                if keyword in article_text:
                    return False, keyword
            return True, "VALID"
        except:
            return True, "UNKNOWN"

    def get_article_date(self, article_element):
        try:
            date_elements = article_element.find_elements(By.CSS_SELECTOR, "[class*='italic'], .label--italic")
            for elem in date_elements:
                date_text = elem.text.strip()
                if date_text:
                    return self.parse_date(date_text)
            return None
        except:
            return None

    def check_keywords_in_article(self, url):
        if url in self.visited_urls:
            return None, None
        self.visited_urls.add(url)
        self.driver.execute_script("window.open('');")
        self.driver.switch_to.window(self.driver.window_handles[-1])
        try:
            self.driver.get(url)
            time.sleep(2)
            page_text = self.driver.find_element(By.TAG_NAME, 'body').text.upper()
            keywords_found = [kw.capitalize() for kw in ['CHINA', 'TAIWAN', 'TAIPEI'] if kw in page_text]
            has_download = bool(self.driver.find_elements(By.XPATH,
                "//*[contains(translate(text(), 'DOWNLOAD', 'download'), 'download')]"))
            return keywords_found, has_download
        except:
            return None, None
        finally:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])

    def scrape_search_page(self, keyword, start_dt, end_dt):
        url = f"https://www.iiss.org/search?query={keyword}"
        self.driver.get(url)
        time.sleep(2)
        self.click_newest_sort()
        page = 1
        self.consecutive_old_articles = 0

        while True:
            print(f"\n處理第 {page} 頁...")
            try:
                articles = self.driver.find_elements(By.CSS_SELECTOR, "a[class*='feature']")
                if not articles:
                    break

                article_data_list = []
                should_stop = False

                for article in articles:
                    try:
                        href = article.get_attribute('href')
                        if not href or '/podcasts/' in href or href in self.visited_urls:
                            continue
                        is_valid, _ = self.check_article_type(article)
                        if not is_valid:
                            continue
                        article_date = self.get_article_date(article)
                        if article_date and article_date < start_dt:
                            self.consecutive_old_articles += 1
                            if self.consecutive_old_articles >= 3:
                                should_stop = True
                            continue
                        elif article_date and article_date > end_dt:
                            continue
                        self.consecutive_old_articles = 0
                        try:
                            raw_title = article.text.strip()
                            lines = [l.strip() for l in raw_title.split('\n')
                                     if l.strip() and not any(m in l for m in
                                     ['January','February','March','April','May','June',
                                      'July','August','September','October','November','December'])]
                            clean_title = ' '.join(lines) if lines else raw_title
                        except:
                            clean_title = "Unknown"
                        article_data_list.append({'url': href, 'clean_title': clean_title, 'date': article_date})
                    except:
                        continue

                for item in article_data_list:
                    keywords, has_download = self.check_keywords_in_article(item['url'])
                    if keywords and has_download:
                        date_str = item['date'].strftime('%Y/%m/%d') if item['date'] else "Unknown"
                        self.articles_data.append({
                            '機構': 'International Institute for Strategic Studies',
                            '日期': date_str,
                            '標題': item['clean_title'],
                            '網址': item['url'],
                            '關鍵字': ', '.join(keywords)
                        })
                        print(f"    ✓ 收錄: {item['clean_title'][:50]}")

                if should_stop:
                    return

                try:
                    next_button = self.driver.find_element(By.CSS_SELECTOR,
                        "a.pagination__next, a[class*='pagination'][class*='next']")
                    self.driver.execute_script("arguments[0].click();", next_button)
                    time.sleep(3)
                    page += 1
                except:
                    break
            except Exception as e:
                print(f"頁面錯誤: {e}")
                break

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("International Institute for Strategic Studies 文章爬蟲程式")
        print("=" * 60)
        self.driver = make_driver()
        try:
            for keyword in ['china', 'taiwan', 'taipei']:
                print(f"\n{'-' * 20}\n搜尋: {keyword.upper()}\n{'-' * 20}")
                self.consecutive_old_articles = 0
                self.scrape_search_page(keyword, start_dt, end_dt)
        except Exception as e:
            print(f"❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.articles_data


# ============================================================
# EVC 爬蟲
# ============================================================

class EVCScraper:
    def __init__(self):
        self.driver = None
        self.articles_data = []
        self.keywords = ["China", "Taiwan", "Taipei"]
        self.institution = "European Value Center for Security Policy"
        self.base_url = "https://europeanvalues.cz/en/red-watch-program-reports/"

    def parse_date_text(self, text):
        text = text.strip()
        for fmt in ["%d %B, %Y", "%d %B %Y", "%B %d, %Y", "%d %b, %Y", "%d %b %Y"]:
            try:
                return datetime.strptime(text, fmt)
            except:
                pass
        m = re.match(r"(\d+)\s+(\w+),?\s+(\d{4})", text)
        if m:
            try:
                return datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %B %Y")
            except:
                try:
                    return datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %b %Y")
                except:
                    pass
        return None

    def check_keywords_in_article(self):
        body_text = ""
        try:
            h1 = self.driver.find_element(By.CSS_SELECTOR, "h1.elementor-heading-title")
            body_text += " " + h1.text
        except:
            pass
        try:
            for sec in self.driver.find_elements(By.CSS_SELECTOR,
                ".elementor-section-boxed .elementor-widget-text-editor"):
                body_text += " " + sec.text
        except:
            pass
        if len(body_text.strip()) < 50:
            try:
                body_text += " " + self.driver.find_element(By.TAG_NAME, "body").text
            except:
                pass
        return [kw for kw in self.keywords if re.search(r'\b' + kw + r'\b', body_text, re.IGNORECASE)]

    def check_download_button(self):
        try:
            for btn in self.driver.find_elements(By.CSS_SELECTOR,
                "a.elementor-button, .elementor-button-wrapper a, a[href*='.pdf']"):
                href = btn.get_attribute("href") or ""
                if ".pdf" in href.lower():
                    return True
        except:
            pass
        return False

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("European Value Center for Security Policy 網站爬蟲")
        print("=" * 60)
        self.driver = make_driver()
        try:
            self.driver.get(self.base_url)
            time.sleep(3)
            page = 1
            consecutive_old = 0
            stop = False

            while not stop:
                print(f"─── 第 {page} 頁 ───")
                time.sleep(2)
                articles = self.driver.find_elements(By.CSS_SELECTOR, "li.elementor-repeater-item")
                if not articles:
                    articles = self.driver.find_elements(By.CSS_SELECTOR,
                        ".elementor-posts-container .elementor-post")
                if not articles:
                    break

                page_items = []
                for art in articles:
                    try:
                        date_el = art.find_element(By.CSS_SELECTOR, "time")
                        art_date = self.parse_date_text(date_el.text.strip())
                        link_el = art.find_element(By.CSS_SELECTOR, "a[href]")
                        href = link_el.get_attribute("href")
                        page_items.append((art_date, href))
                    except:
                        continue

                for art_date, href in page_items:
                    if art_date is None:
                        continue
                    if art_date > end_dt:
                        continue
                    if art_date < start_dt:
                        consecutive_old += 1
                        if consecutive_old >= 3:
                            stop = True
                            break
                        continue
                    consecutive_old = 0
                    self.driver.execute_script("window.open(arguments[0]);", href)
                    self.driver.switch_to.window(self.driver.window_handles[-1])
                    time.sleep(2)
                    try:
                        title = self.driver.find_element(By.CSS_SELECTOR, "h1.elementor-heading-title").text.strip()
                    except:
                        title = href
                    kw_found = self.check_keywords_in_article()
                    has_dl = self.check_download_button()
                    if kw_found and has_dl:
                        self.articles_data.append({
                            '機構': self.institution, '日期': art_date.strftime("%Y/%m/%d"),
                            '標題': title, '網址': href, '關鍵字': ", ".join(kw_found)
                        })
                        print(f"     ✅ 收錄: {title[:50]}")
                    self.driver.close()
                    self.driver.switch_to.window(self.driver.window_handles[0])
                    time.sleep(1)

                if stop:
                    break
                try:
                    nxt = self.driver.find_element(By.CSS_SELECTOR,
                        "a.next.page-numbers, .elementor-pagination a.next, a[aria-label='Next Page']")
                    nxt.click()
                    page += 1
                    time.sleep(3)
                except:
                    break
        except Exception as e:
            print(f"\n❗ 發生錯誤: {e}")
        finally:
            self.driver.quit()
        return self.articles_data


# ============================================================
# Sinopsis 爬蟲
# ============================================================

class SinopsisScraper:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.articles_data = []
        self.keywords = ["China", "Taiwan", "Taipei"]
        self.institution = "sinopsis"
        self.base_url = "https://sinopsis.cz/en/page/{}/"
        self.month_map = {
            "January":1,"February":2,"March":3,"April":4,"May":5,"June":6,
            "July":7,"August":8,"September":9,"October":10,"November":11,"December":12,
        }

    def parse_date(self, text):
        text = text.strip()
        m = re.search(r"(\d{1,2})\.(\d{1,2})\.(\d{4})", text)
        if m:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        m = re.search(r"(\d{1,2})\.\s*([A-Za-z]+)\s+(\d{4})", text)
        if m:
            month = self.month_map.get(m.group(2).capitalize())
            if month:
                return datetime(int(m.group(3)), month, int(m.group(1)))
        m = re.search(r"([A-Za-z]+)\s+(\d{1,2}),?\s+(\d{4})", text)
        if m:
            month = self.month_map.get(m.group(1).capitalize())
            if month:
                return datetime(int(m.group(3)), month, int(m.group(2)))
        return None

    def scrape_article(self, url, start_dt, end_dt):
        self.driver.get(url)
        time.sleep(0.8)
        article_date = None
        for sel in ["span.entry-meta-date a", "span.entry-meta-date", "time.entry-date"]:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, sel)
                article_date = self.parse_date(el.text or el.get_attribute("datetime") or "")
                if article_date:
                    break
            except:
                pass

        if not article_date:
            return None

        if article_date > end_dt or article_date < start_dt:
            return {"date": article_date, "_date_only": True}

        title = ""
        for sel in ["h1.entry-title", "h1", ".entry-title"]:
            try:
                title = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                if title:
                    break
            except:
                pass

        full_text = title
        try:
            full_text += " " + self.driver.find_element(By.CSS_SELECTOR, "div.entry-content").text
        except:
            try:
                full_text += " " + self.driver.find_element(By.TAG_NAME, "body").text
            except:
                pass

        found_kws = [kw for kw in self.keywords if kw in full_text]

        has_pdf = False
        try:
            content_el = self.driver.find_element(By.CSS_SELECTOR, "div.entry-content")
            strong_pdf = content_el.find_elements(By.CSS_SELECTOR, "strong a[href*='.pdf']")
            p_pdf = [lk for lk in content_el.find_elements(By.CSS_SELECTOR, "p a[href*='sinopsis.cz'][href*='.pdf']")
                     if lk.text.strip().lower() == "pdf" or "full text" in lk.find_element(By.XPATH, "..").text.lower()]
            if strong_pdf or p_pdf:
                has_pdf = True
        except:
            pass

        if not found_kws or not has_pdf:
            return {"date": article_date, "_date_only": True}

        return {
            "date": article_date,
            "date_str": article_date.strftime("%Y/%m/%d"),
            "title": title,
            "url": url,
            "keywords": ", ".join(found_kws),
        }

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("Sinopsis.cz 文章爬蟲")
        print("=" * 60)
        self.driver = make_driver()
        self.wait = WebDriverWait(self.driver, 15)
        try:
            page = 1
            stop_scraping = False
            while not stop_scraping:
                list_url = "https://sinopsis.cz/en/" if page == 1 else self.base_url.format(page)
                print(f"\n📄 第 {page} 頁")
                self.driver.get(list_url)
                time.sleep(1.5)
                article_links = []
                try:
                    self.wait.until(EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "h3.entry-title a, h2.entry-title a")))
                    link_els = self.driver.find_elements(By.CSS_SELECTOR, "h3.entry-title a, h2.entry-title a")
                    article_links = [el.get_attribute("href") for el in link_els if el.get_attribute("href")]
                except:
                    break

                if not article_links:
                    break

                consecutive_early = 0
                for link in article_links:
                    if stop_scraping:
                        break
                    try:
                        res = self.scrape_article(link, start_dt, end_dt)
                    except Exception as e:
                        continue
                    if res is None:
                        continue
                    if res.get("_date_only") and res["date"] > end_dt:
                        consecutive_early = 0
                        continue
                    if res.get("_date_only") and res["date"] < start_dt:
                        consecutive_early += 1
                        if consecutive_early >= 3:
                            stop_scraping = True
                            break
                        continue
                    if res.get("_date_only"):
                        consecutive_early = 0
                        continue
                    consecutive_early = 0
                    self.articles_data.append({
                        '機構': self.institution, '日期': res["date_str"],
                        '標題': res["title"], '網址': res["url"], '關鍵字': res["keywords"]
                    })
                    print(f"    ✅ 收錄: {res['title'][:50]}")
                page += 1
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            self.driver.quit()
        return self.articles_data


# ============================================================
# ERIA 爬蟲
# ============================================================

ERIA_MONTHS = {
    "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
    "july":7,"august":8,"september":9,"october":10,"november":11,"december":12
}

class ERIAScraper:
    def __init__(self):
        self.driver = None
        self.wait = None
        self.articles_data = []
        self.keywords = ["China", "Taiwan", "Taipei"]
        self.institution = "Economic Research Institute for ASEAN and East Asia"
        self.base_url = "https://www.eria.org/publications"

    def parse_article_date(self, text):
        if not text:
            return None
        text = text.strip()
        m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", text)
        if m:
            day = int(m.group(1))
            month = ERIA_MONTHS.get(m.group(2).lower())
            year = int(m.group(3))
            if month:
                try:
                    return datetime(year, month, day)
                except:
                    pass
        return None

    def get_article_links_from_grid(self):
        links = []
        seen = set()
        grid_divs = self.driver.find_elements(By.CSS_SELECTOR, "div.grid")
        target_grids = []
        for div in grid_divs:
            try:
                cls = div.get_attribute("class") or ""
                if "gap-6" in cls and "md:grid-cols-2" in cls:
                    target_grids.append(div)
            except:
                continue

        if not target_grids:
            anchors = self.driver.find_elements(By.CSS_SELECTOR, "main a[href]")
            for a in anchors:
                try:
                    href = a.get_attribute("href") or ""
                    if re.search(r"/publications/[^?#/]+$", href) and "/category/" not in href and href not in seen:
                        seen.add(href)
                        links.append(href)
                except:
                    continue
            return links

        for grid in target_grids:
            try:
                for a in grid.find_elements(By.CSS_SELECTOR, "a[href]"):
                    try:
                        href = a.get_attribute("href") or ""
                        if re.search(r"/publications/[^?#/]+$", href) and "/category/" not in href and href not in seen:
                            seen.add(href)
                            links.append(href)
                    except:
                        continue
            except:
                continue
        return links

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("Economic Research Institute for ASEAN and East Asia 文章爬蟲")
        print("=" * 60)
        self.driver = make_driver()
        self.wait = WebDriverWait(self.driver, 20)
        processed_urls = set()
        consecutive_early = 0
        stop_flag = False
        current_limit = 6

        try:
            self.driver.get(self.base_url)
            time.sleep(4)

            while not stop_flag:
                try:
                    self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.grid")))
                except:
                    pass

                article_links = self.get_article_links_from_grid()
                new_links = [l for l in article_links if l not in processed_urls]
                print(f"  本輪新文章：{len(new_links)} 篇")

                if not new_links:
                    break

                for article_url in new_links:
                    if stop_flag:
                        break
                    processed_urls.add(article_url)
                    try:
                        self.driver.get(article_url)
                        time.sleep(2.5)

                        article_date = None
                        try:
                            date_text = self.driver.find_element(By.TAG_NAME, "time").text.strip()
                            article_date = self.parse_article_date(date_text)
                        except:
                            pass

                        if not article_date:
                            try:
                                body_text = self.driver.find_element(By.TAG_NAME, "body").text
                                m = re.search(r"\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}", body_text)
                                if m:
                                    article_date = self.parse_article_date(m.group(0))
                            except:
                                pass

                        if not article_date:
                            continue

                        if article_date > end_dt:
                            consecutive_early = 0
                            continue
                        if article_date < start_dt:
                            consecutive_early += 1
                            if consecutive_early >= 3:
                                stop_flag = True
                            continue

                        consecutive_early = 0

                        has_pdf = bool(self.driver.find_elements(By.XPATH,
                            "//*[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'DOWNLOAD PDF')]"))
                        if not has_pdf:
                            continue

                        try:
                            body_text = self.driver.find_element(By.TAG_NAME, "body").text
                        except:
                            body_text = ""

                        found_keywords = [kw for kw in self.keywords if kw in body_text]
                        if not found_keywords:
                            continue

                        try:
                            title = self.driver.find_element(By.TAG_NAME, "h1").text.strip()
                        except:
                            title = article_url.rstrip("/").split("/")[-1].replace("-", " ").title()

                        self.articles_data.append({
                            '機構': self.institution, '日期': article_date.strftime("%Y/%m/%d"),
                            '標題': title, '網址': article_url, '關鍵字': ", ".join(found_keywords)
                        })
                        print(f"    ✓ 收錄: {title[:50]}")

                    except Exception as e:
                        continue

                if stop_flag:
                    break

                list_url = f"{self.base_url}?limit={current_limit}"
                self.driver.get(list_url)
                time.sleep(3)
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)

                try:
                    show_more_btn = self.wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'SHOW MORE')]")
                    ))
                    self.driver.execute_script("arguments[0].click();", show_more_btn)
                    current_limit += 6
                    time.sleep(3)
                except:
                    break

        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            self.driver.quit()
        return self.articles_data


# ============================================================
# JETRO 爬蟲
# ============================================================

class JETROScraper:
    def __init__(self):
        self.driver = None
        self.articles_data = []
        self.url_reports_survey = "https://www.jetro.go.jp/en/reports/survey/"
        self.url_white_paper = "https://www.jetro.go.jp/en/reports/white_paper/"
        self.headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        self.keywords = ["china", "taiwan", "taipei"]
        self.keyword_display = {"china": "China", "taiwan": "Taiwan", "taipei": "Taipei"}
        self.institution = "Japan External Trade Organization"

    def init_browser(self):
        self.driver = make_driver()

    def browser_open_url(self, url):
        self.driver.get(url)
        time.sleep(1.5)

    def extract_date_from_text(self, text):
        patterns = [
            (r'((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{1,2},?\s+\d{4})', True),
            (r'(\d{4}[/\-\.]\d{1,2}[/\-\.]\d{1,2})', True),
            (r'((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\.?\s+\d{4})', False),
            (r'(\d{4}[/\-]\d{1,2})(?![/\-\d])', False),
        ]
        for pattern, has_day in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                raw = m.group(1).replace('.', '').strip()
                try:
                    dt = dateparser.parse(raw, default=datetime(2000, 1, 1))
                    if not has_day:
                        dt = dt.replace(day=1)
                    return dt.date(), has_day
                except:
                    pass
        return None, False

    def date_to_str(self, d):
        return d.strftime('%Y/%m/%d') if d else '無法解析'

    def get_pdf_text(self, pdf_url):
        if not HAS_PDFPLUMBER:
            return ""
        try:
            resp = requests.get(pdf_url, headers=self.headers, timeout=60)
            if resp.status_code != 200:
                return ""
            import io as _io
            pdf_bytes = _io.BytesIO(resp.content)
            text = ""
            with pdfplumber.open(pdf_bytes) as pdf:
                for page in pdf.pages:
                    text += (page.extract_text() or "")
            return text
        except:
            return ""

    def find_keywords_in_text(self, text):
        found = []
        lower = text.lower()
        for kw in self.keywords:
            if kw in lower:
                found.append(self.keyword_display[kw])
        return found

    def get_pdf_text_with_browser(self, pdf_url):
        self.browser_open_url(pdf_url)
        return self.get_pdf_text(pdf_url)

    def resolve_final_date(self, page_date, page_date_has_day, full_text):
        if page_date and page_date_has_day:
            return page_date
        pdf_date, pdf_has_day = self.extract_date_from_text(full_text) if full_text else (None, False)
        if pdf_date and pdf_has_day:
            return pdf_date
        if page_date:
            return page_date
        if pdf_date:
            return pdf_date
        return date(2000, 1, 1)

    def get_page_soup(self):
        resp = requests.get(self.url_reports_survey, headers=self.headers, timeout=15)
        resp.encoding = 'utf-8'
        return BeautifulSoup(resp.text, 'html.parser')

    def collect_sections(self, soup):
        result = {"latest_topics": [], "quick_survey": [], "business_conditions": {}, "intl_operations": []}
        main = soup.find(id="mainArea") or soup
        current_h2 = current_h3 = current_h4 = ""

        def parse_li_links(li):
            items = []
            li_text = li.get_text(separator=' ')
            for a in li.find_all('a', href=True):
                href = a['href']
                if not href.endswith('.pdf'):
                    continue
                pdf_url = href if href.startswith('http') else 'https://www.jetro.go.jp' + href
                title = a.get_text(strip=True)
                page_date, page_date_has_day = self.extract_date_from_text(li_text)
                items.append({"title": title, "url": pdf_url, "page_date": page_date, "page_date_has_day": page_date_has_day})
            return items

        for tag in main.find_all(['h2', 'h3', 'h4', 'li']):
            name = tag.name
            if name == 'h2':
                current_h2 = tag.get_text(strip=True)
                current_h3 = current_h4 = ""
            elif name == 'h3':
                current_h3 = tag.get_text(strip=True)
                current_h4 = ""
            elif name == 'h4':
                current_h4 = tag.get_text(strip=True)
            elif name == 'li':
                links = parse_li_links(tag)
                if not links:
                    continue
                h2_lower = current_h2.lower()
                h3_lower = current_h3.lower()
                if 'latest topics' in h2_lower and 'quick business survey' not in h3_lower:
                    for item in links:
                        if not any(x['url'] == item['url'] for x in result['latest_topics']):
                            result['latest_topics'].append(item)
                elif 'latest topics' in h2_lower and 'quick business survey' in h3_lower:
                    for item in links:
                        if not any(x['url'] == item['url'] for x in result['quick_survey']):
                            result['quick_survey'].append(item)
                elif 'survey on business conditions' in h3_lower:
                    fy = current_h4 if current_h4 else 'Unknown'
                    result['business_conditions'].setdefault(fy, [])
                    for item in links:
                        if not any(x['url'] == item['url'] for x in result['business_conditions'][fy]):
                            result['business_conditions'][fy].append(item)
                elif 'international operations' in h3_lower:
                    for item in links:
                        if not any(x['url'] == item['url'] for x in result['intl_operations']):
                            result['intl_operations'].append(item)
        return result

    def process_latest_topics(self, items, start_d, end_d):
        collected = []
        consecutive_old = 0
        for item in items:
            title, url, page_date = item['title'], item['url'], item['page_date']
            final_date = page_date if page_date else date(2000, 1, 1)
            in_range = start_d <= final_date <= end_d
            if not in_range and final_date < start_d:
                consecutive_old += 1
                if consecutive_old >= 3:
                    break
                continue
            else:
                consecutive_old = 0
            if not in_range:
                continue
            full_text = self.get_pdf_text_with_browser(url)
            keywords = self.find_keywords_in_text(full_text)
            if not keywords:
                continue
            collected.append({'機構': self.institution, '日期': self.date_to_str(final_date),
                               '標題': title, '網址': url, '關鍵字': ", ".join(keywords)})
        return collected

    def process_quick_survey(self, items, start_d, end_d):
        collected = []
        for item in items:
            title, url, page_date = item['title'], item['url'], item['page_date']
            final_date = page_date if page_date else date(2000, 1, 1)
            if not (start_d <= final_date <= end_d):
                break
            full_text = self.get_pdf_text_with_browser(url)
            keywords = self.find_keywords_in_text(full_text)
            if not keywords:
                continue
            collected.append({'機構': self.institution, '日期': self.date_to_str(final_date),
                               '標題': title, '網址': url, '關鍵字': ", ".join(keywords)})
        return collected

    def process_business_conditions(self, bc_dict, start_d, end_d):
        if not bc_dict:
            return []
        latest_fy = sorted(bc_dict.keys(), reverse=True)[0]
        items = bc_dict[latest_fy]
        collected = []
        for item in items:
            raw_title, url = item['title'], item['url']
            page_date, page_date_has_day = item['page_date'], item['page_date_has_day']
            clean_title = re.sub(r'\s*\([\d.]+\s*(?:MB|KB)\)', '', raw_title).strip()
            title = f"({latest_fy})Survey on Business Conditions of Japanese-Affiliated Companies Overseas: {clean_title}"
            full_text = self.get_pdf_text_with_browser(url)
            final_date = self.resolve_final_date(page_date, page_date_has_day, full_text)
            keywords = self.find_keywords_in_text(full_text)
            if not (start_d <= final_date <= end_d) or not keywords:
                continue
            collected.append({'機構': self.institution, '日期': self.date_to_str(final_date),
                               '標題': title, '網址': url, '關鍵字': ", ".join(keywords)})
        return collected

    def process_intl_operations(self, items, start_d, end_d):
        if not items:
            return []
        item = items[0]
        raw_title, url, page_date = item['title'], item['url'], item['page_date']
        fy_match = re.search(r'FY\s*(\d{4})', raw_title, re.IGNORECASE)
        fy_label = f"FY {fy_match.group(1)}" if fy_match else raw_title
        title = f"({fy_label}) Survey on the International Operations of Japanese Firms"
        final_date = page_date if page_date else date(2000, 1, 1)
        if not (start_d <= final_date <= end_d):
            return []
        full_text = self.get_pdf_text_with_browser(url)
        keywords = self.find_keywords_in_text(full_text)
        if not keywords:
            return []
        return [{'機構': self.institution, '日期': self.date_to_str(final_date),
                 '標題': title, '網址': url, '關鍵字': ", ".join(keywords)}]

    def collect_white_paper(self, soup):
        main = soup.find(id="mainArea") or soup
        results = []
        for li in main.find_all('li'):
            pdfs = []
            for a in li.find_all('a', href=True):
                if a['href'].endswith('.pdf'):
                    pdf_url = a['href'] if a['href'].startswith('http') else 'https://www.jetro.go.jp' + a['href']
                    pdfs.append({"title_raw": a.get_text(strip=True), "url": pdf_url})
            if pdfs:
                li_text = li.get_text(separator=' ')
                year_match = re.search(r'\b(20\d{2})\b', li_text)
                year = int(year_match.group(1)) if year_match else None
                for p in pdfs:
                    p['year'] = year
                results = pdfs
                break
        return results

    def get_white_paper_date(self, items, year):
        fallback = date(year, 7, 30)
        if len(items) == 1:
            text = self.get_pdf_text_with_browser(items[0]['url'])
            d, has_day = self.extract_date_from_text(text)
            if d:
                return d if has_day else date(d.year, d.month, 1)
            return fallback
        text2 = self.get_pdf_text_with_browser(items[1]['url'])
        d2, has_day2 = self.extract_date_from_text(text2)
        if d2 and has_day2:
            return d2
        text1 = self.get_pdf_text_with_browser(items[0]['url'])
        d1, has_day1 = self.extract_date_from_text(text1)
        if d1:
            return date(d1.year, d1.month, 1)
        return fallback

    def process_white_paper(self, start_d, end_d):
        print("\n" + "-"*20 + "\n處理 JETRO White Paper...\n" + "-"*20)
        self.browser_open_url(self.url_white_paper)
        resp = requests.get(self.url_white_paper, headers=self.headers, timeout=15)
        resp.encoding = 'utf-8'
        soup = BeautifulSoup(resp.text, 'html.parser')
        items = self.collect_white_paper(soup)
        if not items:
            return []
        year = items[0]['year']
        final_date = self.get_white_paper_date(items, year)
        if not (start_d <= final_date <= end_d):
            return []
        collected = []
        for i, item in enumerate(items):
            raw = item['title_raw']
            url = item['url']
            title = f"JETRO Global Trade and Investment Report {year}" + ("-Overview" if 'overview' in raw.lower() else "")
            full_text = self.get_pdf_text(url)
            keywords = self.find_keywords_in_text(full_text)
            if not keywords:
                continue
            collected.append({'機構': self.institution, '日期': self.date_to_str(final_date),
                               '標題': title, '網址': url, '關鍵字': ", ".join(keywords)})
        return collected

    def run(self, start_dt, end_dt):
        print("\n" + "="*60 + "\nJapan External Trade Organization 文章爬蟲\n" + "="*60)
        start_d = start_dt.date() if isinstance(start_dt, datetime) else start_dt
        end_d   = end_dt.date()   if isinstance(end_dt,   datetime) else end_dt

        soup = self.get_page_soup()
        sections = self.collect_sections(soup)
        self.init_browser()
        self.browser_open_url(self.url_reports_survey)

        try:
            self.articles_data += self.process_white_paper(start_d, end_d)
            self.articles_data += self.process_latest_topics(sections['latest_topics'], start_d, end_d)
            self.articles_data += self.process_quick_survey(sections['quick_survey'], start_d, end_d)
            self.articles_data += self.process_business_conditions(sections['business_conditions'], start_d, end_d)
            self.articles_data += self.process_intl_operations(sections['intl_operations'], start_d, end_d)
        except Exception as e:
            print(f"\n❌ 錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        return self.articles_data


# ============================================================
# 主程式
# ============================================================

def main():
    all_articles = []

    scrapers = [
        ("IFRI",    lambda: IFRIScraper().run(start_date, end_date)),
        ("ECFR",    lambda: ECFRScraper().run(start_date, end_date)),
        ("IDS",     lambda: IDSScraper().run(start_date, end_date)),
        ("IISS",    lambda: IISSScraper().run(start_date, end_date)),
        ("EVC",     lambda: EVCScraper().run(start_date, end_date)),
        ("Sinopsis",lambda: SinopsisScraper().run(start_date, end_date)),
        ("ERIA",    lambda: ERIAScraper().run(start_date, end_date)),
        ("JETRO",   lambda: JETROScraper().run(start_date, end_date)),
    ]

    for name, run_fn in scrapers:
        try:
            articles = run_fn()
            all_articles.extend(articles)
            print(f"  ✓ {name} 完成，本次收錄 {len(articles)} 篇")
        except Exception as e:
            print(f"  ❌ {name} 發生未預期錯誤，跳過：{e}")
            traceback.print_exc()

    excel_path = "combined_articles_G5_normal.xlsx"
    total_count = save_to_excel(all_articles, excel_path)

    print("\n📧 正在寄送 Email...")
    log_text = logger.get_log()
    send_email(excel_path, log_text, total_count)

    print("🎉 全部完成！")


if __name__ == "__main__":
    main()
