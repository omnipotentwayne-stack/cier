#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合併爬蟲：KIEP + DIIS + IAI + NUPI + ORF + IFANS + ISPI
自動計算日期區間（1/11/21號），headless模式，執行完寄信
"""

import re
import io
import sys
import time
import calendar
import os
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import openpyxl
from dateutil import parser as dateutil_parser

try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False


# ═══════════════════════════════════════════════════════════════════════
# LOG 攔截：把所有 print 同時寫到畫面 + 記憶體
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
    msg["Subject"] = f"G2爬蟲結果 {today.strftime('%Y/%m/%d')}（區間 {start_date.date()} ~ {end_date.date()}）"

    body = (
        f"G2爬蟲執行完畢。\n"
        f"執行日期：{today.strftime('%Y/%m/%d %H:%M')}\n"
        f"資料區間：{start_date.date()} ~ {end_date.date()}\n"
        f"共收錄：{total_count} 筆\n\n"
        f"詳細 log 請見附件 G2_log.txt，Excel 結果請見附件。"
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
    log_part.add_header("Content-Disposition", "attachment; filename=G2_log.txt")
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
    opts = webdriver.ChromeOptions()
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


# ══════════════════════════════════════════════════════════════════════
# KIEP Crawler
# ══════════════════════════════════════════════════════════════════════

class KIEPCrawler:
    def __init__(self):
        self.driver = None
        self.keywords = ['China', 'Taiwan', 'Taipei']
        self.institution = "Korea Institute for International Economic Policy"
        self.results = []
        self.base_url = "https://www.kiep.go.kr"

    def setup_driver(self):
        opts = build_chrome_options()
        opts.add_argument('--lang=en-US')
        if USE_WEBDRIVER_MANAGER:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=opts)
        else:
            self.driver = webdriver.Chrome(options=opts)
        self.driver.implicitly_wait(10)

    def parse_article_date(self, date_text):
        if not date_text:
            return None
        date_text = date_text.strip()
        patterns = [
            (r'(\d{4})\.(\d{1,2})\.(\d{1,2})', 'ymd'),
            (r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', 'ymd'),
            (r'(\d{1,2})\.(\d{1,2})\.(\d{4})', 'dmy'),
        ]
        for pattern, fmt in patterns:
            m = re.search(pattern, date_text)
            if m:
                g = m.groups()
                try:
                    if fmt == 'ymd':
                        return datetime(int(g[0]), int(g[1]), int(g[2]))
                    elif fmt == 'dmy':
                        return datetime(int(g[2]), int(g[1]), int(g[0]))
                except ValueError:
                    continue
        return None

    def check_keywords(self, text):
        if not text:
            return []
        text_lower = text.lower()
        return list({kw for kw in self.keywords if kw.lower() in text_lower})

    def get_article_list_items(self):
        try:
            ul = self.driver.find_element(By.CSS_SELECTOR, "ul.board_book")
            items = ul.find_elements(By.TAG_NAME, "li")
            if items:
                return items, 'board_book'
        except NoSuchElementException:
            pass
        try:
            ul = self.driver.find_element(By.CSS_SELECTOR, "ul.board_report")
            items = ul.find_elements(By.TAG_NAME, "li")
            if items:
                return items, 'board_report'
        except NoSuchElementException:
            pass
        print("  ⚠ 未找到文章列表元素")
        return [], None

    def extract_list_item_report(self, li):
        try:
            try:
                number_elem = li.find_element(By.CSS_SELECTOR, "b.number")
                article_number = number_elem.text.strip()
            except:
                article_number = ""
            link_elem = li.find_element(By.TAG_NAME, "a")
            title = link_elem.text.strip()
            href  = link_elem.get_attribute('href')
            article_url = href if href and href.startswith('http') else self.base_url + (href or '')
            return {
                'number': article_number, 'title': title, 'url': article_url,
                'date': None, 'valid': bool(title and article_url and 'javascript' not in article_url.lower())
            }
        except Exception:
            return {'valid': False}

    def extract_list_item_book(self, li):
        try:
            desc = li.find_element(By.CSS_SELECTOR, "div.desc")
            a_elem = desc.find_element(By.CSS_SELECTOR, "a")
            href = a_elem.get_attribute('href') or ''
            article_url = href if href.startswith('http') else self.base_url + href
            try:
                title = a_elem.find_element(By.CSS_SELECTOR, "strong.title").text.strip()
            except:
                title = a_elem.text.strip()
            date = None
            author = ''
            try:
                info_p = desc.find_element(By.CSS_SELECTOR, "p.info")
                spans  = info_p.find_elements(By.TAG_NAME, "span")
                for i, span in enumerate(spans):
                    text = span.text.strip()
                    if i == 0 and text and not re.search(r'\d{4}', text):
                        author = text
                    d = self.parse_article_date(text)
                    if d:
                        date = d
            except:
                pass
            return {
                'number': '', 'title': title, 'url': article_url, 'date': date,
                'author': author, 'valid': bool(title and article_url and 'javascript' not in article_url.lower())
            }
        except Exception:
            return {'valid': False}

    def get_article_details(self, list_type='board_report'):
        result = {'title': '', 'date': None, 'content': '', 'author': ''}
        try:
            time.sleep(2)
            if list_type == 'board_book':
                for sel in ["div.board_view strong.title", "div.board_book strong.title",
                            "div.desc strong.title", "div.desc strong"]:
                    try:
                        text = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                        if text:
                            result['title'] = text
                            break
                    except:
                        pass
                try:
                    info_p = self.driver.find_element(By.CSS_SELECTOR, "p.info")
                    spans  = info_p.find_elements(By.TAG_NAME, "span")
                    for i, span in enumerate(spans):
                        text = span.text.strip()
                        if i == 0 and text and not re.search(r'\d{4}', text):
                            result['author'] = text
                        d = self.parse_article_date(text)
                        if d:
                            result['date'] = d
                except:
                    pass
                for sel in ["div.cont", "div.txt", "div.view_cont", "div.board_view", "div.contents"]:
                    try:
                        elem = self.driver.find_element(By.CSS_SELECTOR, sel)
                        text = elem.text.strip()
                        if text:
                            result['content'] = text
                            break
                    except:
                        pass
            else:
                for sel in ["div.subject", "h1.subject", "h2.subject",
                            "div.board_view .subject", ".view_title"]:
                    try:
                        result['title'] = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                        if result['title']:
                            break
                    except:
                        pass
                for sel in ["div.board_info", "div.view_info", ".info",
                            "div.board_view .info", "p.info"]:
                    try:
                        info_elem = self.driver.find_element(By.CSS_SELECTOR, sel)
                        info_text = info_elem.text.strip()
                        m = re.search(r'^([^D]+?)(?=\s+Date)', info_text)
                        if m:
                            result['author'] = m.group(1).strip()
                        result['date'] = self.parse_article_date(info_text)
                        if result['date']:
                            break
                    except:
                        pass
                if not result['date']:
                    for sel in ["span.date", "div.date", "td.date", ".wdate"]:
                        try:
                            result['date'] = self.parse_article_date(
                                self.driver.find_element(By.CSS_SELECTOR, sel).text)
                            if result['date']:
                                break
                        except:
                            pass
                for sel in ["div.board_view", "div.view_content", "div.contents",
                            "div.content", ".article_content"]:
                    try:
                        elem = self.driver.find_element(By.CSS_SELECTOR, sel)
                        text = elem.text.strip()
                        if text:
                            result['content'] = text
                            break
                    except:
                        pass
        except Exception as e:
            print(f"    获取详情出错: {e}")
        return result

    def crawl_page(self, url, start_dt, end_dt):
        print(f"\n{'-' * 20}")
        print(f"開始爬取: {url}")
        print(f"{'-' * 20}")
        self.driver.get(url)
        time.sleep(3)
        consecutive_early = 0
        page_num = 1
        while True:
            print(f"\n--- 第 {page_num} 页 ---")
            try:
                articles, list_type = self.get_article_list_items()
                if not articles:
                    print("⚠ 页面无文章，停止爬取该分类")
                    break
                print(f"找到 {len(articles)} 篇文章  [结构: {list_type}]")
                for idx, li in enumerate(articles, 1):
                    try:
                        if list_type == 'board_book':
                            info = self.extract_list_item_book(li)
                        else:
                            info = self.extract_list_item_report(li)
                        if not info['valid']:
                            continue
                        title       = info['title']
                        article_url = info['url']
                        number      = info.get('number', '')
                        list_date   = info.get('date')
                        label = f"#{number} " if number else ""
                        print(f"\n[{idx}/{len(articles)}] {label}{title[:55]}...")
                        main_window = self.driver.current_window_handle
                        self.driver.execute_script("window.open(arguments[0]);", article_url)
                        time.sleep(1)
                        self.driver.switch_to.window(self.driver.window_handles[-1])
                        time.sleep(2)
                        details = self.get_article_details(list_type)
                        if not details['title']:
                            details['title'] = title
                        if not details['date'] and list_date:
                            details['date'] = list_date
                        if not details['date']:
                            print("  ⚠ 无法解析日期，跳过")
                            self.driver.close()
                            self.driver.switch_to.window(main_window)
                            continue
                        print(f"  📅 日期: {details['date'].strftime('%Y/%m/%d')}")
                        if details['date'] > end_dt:
                            print(f"  ⏭  晚于结束日期")
                            consecutive_early = 0
                            self.driver.close()
                            self.driver.switch_to.window(main_window)
                            continue
                        if details['date'] < start_dt:
                            consecutive_early += 1
                            print(f"  ⏭  早于起始日期 (连续: {consecutive_early})")
                            self.driver.close()
                            self.driver.switch_to.window(main_window)
                            if consecutive_early >= 3:
                                print("⏹ 连续3篇文章早于起始日期，停止搜索该分类")
                                return
                            continue
                        consecutive_early = 0
                        found = list(set(
                            self.check_keywords(details['title']) +
                            self.check_keywords(details['content'])
                        ))
                        if found:
                            print(f"  ✓ 找到关键字: {', '.join(found)}")
                            self.results.append({
                                '機構': self.institution,
                                '日期': details['date'].strftime('%Y/%m/%d'),
                                '標題': details['title'],
                                '網址': article_url,
                                '關鍵字': ', '.join(sorted(found))
                            })
                        else:
                            print(f"  ✗ 无关键字匹配")
                        self.driver.close()
                        self.driver.switch_to.window(main_window)
                        time.sleep(0.5)
                    except Exception as e:
                        print(f"  ❌ 处理文章出错: {e}")
                        try:
                            if len(self.driver.window_handles) > 1:
                                self.driver.close()
                            self.driver.switch_to.window(self.driver.window_handles[0])
                        except:
                            pass
                        continue
                next_clicked = False
                for sel in ["a.next:not(.disabled)", "a[title='next page']",
                            "a.page_next", "div.paging a.next", "div.board_paging a.next"]:
                    try:
                        btn = self.driver.find_element(By.CSS_SELECTOR, sel)
                        if btn.is_displayed() and btn.is_enabled():
                            cls = btn.get_attribute('class') or ''
                            if 'disabled' not in cls.lower():
                                btn.click()
                                next_clicked = True
                                print(f"\n▶ 进入第 {page_num + 1} 页...")
                                time.sleep(3)
                                page_num += 1
                                break
                    except:
                        continue
                if not next_clicked:
                    print("\n📄 已到最后一页或未找到下一页按钮")
                    break
            except Exception as e:
                print(f"❌ 页面处理出错: {e}")
                break

    def run(self, start_dt, end_dt):
        print("\n" + "=" * 60)
        print("KIEP網站爬蟲")
        print("=" * 60)
        print("🔧 初始化浏览器...")
        self.setup_driver()
        print("✓ 浏览器启动成功\n")
        urls = [
            'https://www.kiep.go.kr/gallery.es?mid=a20205030000&bid=0001',
            'https://www.kiep.go.kr/gallery.es?mid=a20205010000&bid=0007',
            'https://www.kiep.go.kr/gallery.es?mid=a20205020000&bid=0008',
        ]
        labels = ['Working Papers', 'World Economy Brief', 'KIEP Opinions']
        for i, (url, label) in enumerate(zip(urls, labels), 1):
            print(f"\n[{i}/{len(urls)}] {label}")
            self.crawl_page(url, start_dt, end_dt)
        print("\n🎉 KIEP 爬取完成！")
        if self.driver:
            self.driver.quit()
            print("✓ 浏览器已关闭\n")


# ══════════════════════════════════════════════════════════════════════
# DIIS Scraper
# ══════════════════════════════════════════════════════════════════════

class DIISScraper:
    def __init__(self):
        self.driver = None
        self.articles = []
        self.institution = "Danish Institute for International Studies"
        self.publication_types = {
            'Article': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A921',
            'DIIS Policy brief': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A916',
            'DIIS Comment': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A915',
            'DIIS Report': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A913',
            'Book': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A912',
            'DIIS Working Paper': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A914',
            'Brief': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A935',
            'Working papers etc.': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A923',
            'Chapter': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A919',
            'Book Chapter': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A934',
            'Report': 'https://www.diis.dk/en/research?s={keyword}&sort_by=d&f%5B0%5D=p%3A924'
        }

    def setup_driver(self):
        try:
            if USE_WEBDRIVER_MANAGER:
                service = Service(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service, options=build_chrome_options())
            else:
                self.driver = webdriver.Chrome(options=build_chrome_options())
            self.driver.implicitly_wait(10)
        except Exception as e:
            print(f"❌ 無法初始化瀏覽器: {e}")
            sys.exit(1)

    def parse_date_from_article(self, article_url):
        full_title = None
        parsed_date = None
        try:
            self.driver.get(article_url)
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)
            for sel in ["//h1[contains(@class,'page-title')]", "//h1[contains(@class,'title')]",
                        "//div[contains(@class,'page-header')]//h1", "//main//h1", "//h1"]:
                try:
                    el = self.driver.find_element(By.XPATH, sel)
                    text = el.text.strip()
                    if text:
                        full_title = text
                        break
                except NoSuchElementException:
                    continue
            date_selectors = [
                "//time[@datetime]", "//span[contains(@class, 'date')]",
                "//div[contains(@class, 'date')]",
                "//div[contains(@class, 'field-name-post-date')]//div[@class='field-item']",
                "//div[contains(@class, 'published')]",
                "//span[contains(@class, 'date-display-single')]",
                "//meta[@property='article:published_time']",
                "//meta[@name='date']", "//meta[@property='og:published_time']"
            ]
            for selector in date_selectors:
                try:
                    if selector.startswith("//meta"):
                        date_element = self.driver.find_element(By.XPATH, selector)
                        date_str = date_element.get_attribute('content')
                    else:
                        date_element = self.driver.find_element(By.XPATH, selector)
                        date_str = (date_element.get_attribute('datetime') or date_element.text)
                    if date_str:
                        date_str = (date_str.strip().split('T')[0] if 'T' in date_str else date_str.strip())
                        for fmt in ['%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y', '%Y/%m/%d',
                                    '%d %B %Y', '%B %d, %Y', '%d-%m-%Y', '%Y.%m.%d',
                                    '%d %b %Y', '%b %d, %Y']:
                            try:
                                candidate = datetime.strptime(date_str, fmt)
                                if 2000 <= candidate.year <= 2030:
                                    parsed_date = candidate
                                    break
                            except ValueError:
                                continue
                    if parsed_date:
                        break
                except NoSuchElementException:
                    continue
        except Exception:
            pass
        return parsed_date, full_title

    def check_pdf_download(self, article_url):
        try:
            for selector in [
                "//a[contains(translate(text(), 'EXTERNAL LINK', 'external link'), 'external link')]",
                "//span[contains(translate(text(), 'EXTERNAL LINK', 'external link'), 'external link')]"
            ]:
                try:
                    if self.driver.find_elements(By.XPATH, selector):
                        return False
                except Exception:
                    continue
            for selector in [
                "//a[contains(translate(text(), 'DOWNLOAD', 'download'), 'download')]",
                "//a[contains(@href, '.pdf')]",
                "//a[contains(@class, 'download')]",
                "//div[contains(@class, 'download')]//a",
                "//button[contains(translate(text(), 'DOWNLOAD', 'download'), 'download')]"
            ]:
                try:
                    for element in self.driver.find_elements(By.XPATH, selector):
                        text = element.text.lower()
                        href = element.get_attribute('href') or ''
                        if ('external' not in text and
                                ('pdf' in text or '.pdf' in href.lower() or 'download' in text)):
                            return True
                except Exception:
                    continue
            return False
        except Exception:
            return False

    def get_article_links(self):
        article_links = []
        try:
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(3)
            for strategy in ["//h2//a[@href] | //h3//a[@href]",
                             "//div[contains(@class, 'field-name-title')]//a[@href]",
                             "//div[contains(@class, 'node-')]//a[@href]",
                             "//div[contains(@class, 'views-row')]//a[@href]"]:
                try:
                    for element in self.driver.find_elements(By.XPATH, strategy):
                        try:
                            href = element.get_attribute('href')
                            if not href or 'javascript:' in href:
                                continue
                            valid_paths = ['/en/research/', '/en/publications/', '/publikation/', '/node/']
                            exclude_paths = ['/about', '/contact', '/privacy', '/en/home',
                                             '/search', '/user/', '/admin/', '/library', '/experts', '/events']
                            if not any(p in href for p in valid_paths):
                                continue
                            if any(ex in href.lower() for ex in exclude_paths):
                                continue
                            title = element.text.strip()
                            if not title:
                                continue
                            exclude_titles = ['diis comment', 'diis report', 'diis policy brief',
                                              'diis working paper', 'article', 'report', 'book',
                                              'working paper', 'policy brief', 'comment', 'brief',
                                              'chapter', 'book chapter', 'working papers etc.', 'read more']
                            if title.isdigit() or title.lower().strip() in exclude_titles:
                                continue
                            if not any(a['url'] == href for a in article_links):
                                article_links.append({'url': href, 'title': title})
                        except Exception:
                            continue
                    if article_links:
                        break
                except Exception:
                    continue
            return article_links
        except Exception:
            return []

    def scrape_page(self, url, keyword, start_dt, end_dt):
        try:
            self.driver.get(url)
            article_links = self.get_article_links()
            if not article_links:
                return 0, False, True
            print(f"      找到 {len(article_links)} 篇")
            consecutive_old = 0
            articles_added = 0
            for idx, article_info in enumerate(article_links, 1):
                article_url = article_info['url']
                list_title = article_info['title']
                article_date, full_title = self.parse_date_from_article(article_url)
                title = full_title if full_title else list_title
                print(f"      [{idx}/{len(article_links)}] {title}", end='')
                if not article_date:
                    print(f" - 無日期")
                    continue
                date_str = article_date.strftime('%Y/%m/%d')
                if article_date < start_dt:
                    print(f" - {date_str} (太早)")
                    consecutive_old += 1
                    if consecutive_old >= 3:
                        print(f"      ⛔ 連續3篇早於起始日期")
                        return articles_added, True, False
                    continue
                elif article_date > end_dt:
                    print(f" - {date_str} (太晚)")
                    consecutive_old = 0
                    continue
                consecutive_old = 0
                has_pdf = self.check_pdf_download(article_url)
                if has_pdf:
                    print(f" - {date_str} ✅")
                    self.articles.append({
                        'institution': self.institution,
                        'date': article_date.strftime('%Y/%m/%d'),
                        'title': title,
                        'url': article_url,
                        'keyword': keyword
                    })
                    articles_added += 1
                else:
                    print(f" - {date_str} (無PDF)")
            return articles_added, False, False
        except Exception as e:
            print(f"      錯誤: {str(e)[:50]}")
            return 0, False, True

    def scrape_publication_type(self, pub_type, base_url, keyword, start_dt, end_dt):
        page = 0
        total_articles = 0
        while True:
            url = base_url if page == 0 else f"{base_url}&page={page}"
            print(f"  {pub_type} p.{page + 1}", end=' ')
            articles_found, should_stop, no_articles = self.scrape_page(url, keyword, start_dt, end_dt)
            total_articles += articles_found
            if no_articles:
                print("(無)")
                break
            if should_stop:
                print()
                break
            try:
                next_buttons = self.driver.find_elements(By.XPATH,
                    "//a[contains(@rel, 'next') or contains(@class, 'next') or "
                    "contains(text(), 'Next') or contains(text(), '›') or "
                    "contains(@class, 'pager-next')]")
                has_next = any('disabled' not in (btn.get_attribute('class') or '') for btn in next_buttons)
                if not has_next:
                    break
            except Exception:
                break
            page += 1
            time.sleep(2)
        if total_articles > 0:
            print(f"    ✓ 共 {total_articles} 篇")
        return total_articles

    def scrape_all_keywords(self, start_dt, end_dt):
        keywords = ['china', 'taiwan', 'taipei']
        total_count = 0
        for keyword in keywords:
            print(f"\n{'-' * 20}")
            print(f"🔍 關鍵字: {keyword.upper()}")
            print(f"{'-' * 20}")
            keyword_capitalized = keyword.capitalize()
            keyword_count = 0
            for pub_type, url_template in self.publication_types.items():
                base_url = url_template.format(keyword=keyword)
                count = self.scrape_publication_type(pub_type, base_url, keyword_capitalized, start_dt, end_dt)
                keyword_count += count
            print(f"\n✅ {keyword.upper()} 總計: {keyword_count} 篇")
            total_count += keyword_count
            time.sleep(3)
        return total_count

    def run(self, start_dt, end_dt):
        print("=" * 60)
        print("DIIS 爬蟲程式")
        print("=" * 60)
        self.setup_driver()
        print("✓ 瀏覽器已啟動")
        self.scrape_all_keywords(start_dt, end_dt)
        print("\n🎉 DIIS 爬取完成！")
        if self.driver:
            self.driver.quit()
            print("✓ 瀏覽器已關閉")


# ══════════════════════════════════════════════════════════════════════
# IAI Scraper
# ══════════════════════════════════════════════════════════════════════

class IAIScraper:
    def __init__(self):
        self.base_urls = {
            'Documenti IAI': 'https://www.iai.it/en/publications/list/documenti-iai',
            'Reports for Parliament': 'https://www.iai.it/en/publications/list/reports-for-parliament',
            'IAI Commentaries': 'https://www.iai.it/en/publications/list/iai-commentaries',
            'IAI Papers': 'https://www.iai.it/en/publications/list/iai-papers',
            'IAI Research Studies': 'https://www.iai.it/en/publications/list/iai-research-studies',
            'IAI Briefs': 'https://www.iai.it/en/publications/list/iai-briefs',
            'Books': 'https://www.iai.it/en/publications/list/books',
            'Other papers and articles': 'https://www.iai.it/en/publications/list/other-papers-and-articles'
        }
        self.keywords = ['China', 'Taiwan', 'Taipei']
        self.results = []

    def parse_article_date(self, date_text):
        try:
            date_text = date_text.strip()
            match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', date_text)
            if match:
                day, month, year = match.groups()
                return datetime(int(year), int(month), int(day))
            return dateutil_parser.parse(date_text, dayfirst=True)
        except Exception:
            return None

    def check_pdf_button(self, driver):
        try:
            if driver.find_elements(By.CSS_SELECTOR, 'a[href*=".pdf"]'):
                return True
            if driver.find_elements(By.XPATH, "//*[contains(text(), 'PDF') or contains(text(), 'pdf')]"):
                return True
            return False
        except Exception:
            return False

    def scrape_category(self, driver, category_name, base_url, keyword, start_dt, end_dt):
        page = 0
        consecutive_old_articles = 0
        total_articles_found = 0
        print(f"\n🔍 開始爬取 {category_name} - 關鍵字: {keyword}")
        while True:
            url = (f"{base_url}?populate={keyword.lower()}&date_year=&tema=All&tag=All" if page == 0
                   else f"{base_url}?populate={keyword.lower()}&date_year=&tema=All&tag=All&page={page}")
            driver.get(url)
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(3)
                article_links = driver.find_elements(By.CSS_SELECTOR, 'h3.card-title a, h5.card-title a')
                if not article_links:
                    article_links = driver.find_elements(By.CSS_SELECTOR, 'a.stretched-link')
                if not article_links:
                    print(f"⚠️ 第 {page + 1} 頁沒有找到文章")
                    break
                article_urls = list(set([link.get_attribute('href') for link in article_links if link.get_attribute('href')]))
                for idx, article_url in enumerate(article_urls, 1):
                    try:
                        driver.get(article_url)
                        time.sleep(2)
                        try:
                            title = driver.find_element(By.CSS_SELECTOR, 'h1.page-title, h1').text.strip()
                        except:
                            title = "無標題"
                        article_date = None
                        for selector in ['div.field--name-field-data', 'div.field--type-datetime',
                                         'span.date', 'time', 'div.field_item']:
                            try:
                                for date_element in driver.find_elements(By.CSS_SELECTOR, selector):
                                    date_text = date_element.text.strip()
                                    if date_text and len(date_text) >= 8:
                                        article_date = self.parse_article_date(date_text)
                                        if article_date:
                                            break
                                if article_date:
                                    break
                            except:
                                continue
                        if not article_date:
                            continue
                        if article_date < start_dt:
                            consecutive_old_articles += 1
                            if consecutive_old_articles >= 10:
                                return True
                            continue
                        if article_date > end_dt:
                            consecutive_old_articles = 0
                            continue
                        consecutive_old_articles = 0
                        if self.check_pdf_button(driver):
                            total_articles_found += 1
                            self.results.append({
                                '機構': 'Istituto Affari Internazionali',
                                '日期': article_date.strftime('%Y/%m/%d'),
                                '標題': title,
                                '網址': article_url,
                                '關鍵字': keyword
                            })
                            print(f"  ✅ 收錄 ({article_date.strftime('%Y/%m/%d')}): {title[:50]}")
                    except Exception as e:
                        continue
                page += 1
            except Exception as e:
                break
        return False

    def run(self, start_dt, end_dt):
        print("=" * 60)
        print("IAI 爬蟲程式")
        print("=" * 60)
        driver = webdriver.Chrome(options=build_chrome_options())
        try:
            for keyword in self.keywords:
                print(f"\n{'-' * 20}")
                print(f"開始爬取關鍵字: {keyword}")
                print(f"{'-' * 20}")
                for category_name, base_url in self.base_urls.items():
                    self.scrape_category(driver, category_name, base_url, keyword, start_dt, end_dt)
                    time.sleep(2)
        except Exception as e:
            print(f"\n❌ 執行時發生錯誤: {str(e)}")
        finally:
            driver.quit()
            print("\n🔚 瀏覽器已關閉")
        print("\n🎉 IAI 爬取完成！")


# ══════════════════════════════════════════════════════════════════════
# NUPI Scraper
# ══════════════════════════════════════════════════════════════════════

class NUPIScraper:
    def __init__(self):
        self.base_url = "https://www.nupi.no/en/publications"
        self.publication_types = [
            "pub_working_paper", "pub_policy_brief", "pub_research_paper",
            "pub_report", "pub_book", "pub_chapter"
        ]
        self.keywords = ["China", "Taiwan", "Taipei"]
        self.results = []

    def parse_date_from_page(self, date_text):
        try:
            return datetime.strptime(date_text.replace("Published:", "").strip(), '%B %d, %Y')
        except:
            return None

    def check_keywords(self, text):
        text_lower = text.lower()
        return [kw for kw in self.keywords if kw.lower() in text_lower]

    def scrape_article(self, driver, article_url, start_dt, end_dt):
        try:
            driver.get(article_url)
            time.sleep(2)
            try:
                title = driver.find_element(By.CSS_SELECTOR, "h1.title").text
            except NoSuchElementException:
                title = "（無標題）"
            try:
                date_text = driver.find_element(By.CSS_SELECTOR, "div.published_date").text
                article_date = self.parse_date_from_page(date_text)
                if not article_date or not (start_dt <= article_date <= end_dt):
                    return None, article_date
            except NoSuchElementException:
                return None, None
            try:
                content = driver.find_element(By.CSS_SELECTOR, "div.publication_content").text
            except NoSuchElementException:
                content = ""
            found_keywords = self.check_keywords(title + " " + content)
            if not found_keywords:
                return None, article_date
            try:
                has_download = any("download" in btn.text.lower()
                                   for btn in driver.find_elements(By.CSS_SELECTOR, "a.component_button.hover"))
                if not has_download:
                    return None, article_date
            except NoSuchElementException:
                return None, article_date
            return {
                "機構": "Norwegian Institute of International Affairs",
                "日期": article_date.strftime('%Y/%m/%d'),
                "標題": title,
                "網址": article_url,
                "關鍵字": ", ".join(found_keywords)
            }, article_date
        except Exception as e:
            return None, None

    def scrape_publication_type(self, driver, pub_type, start_dt, end_dt):
        display_type = pub_type.replace("pub_", "")
        print(f"\n{'-' * 20}")
        print(f"📚 開始爬取類型：{display_type}")
        print(f"{'-' * 20}")
        page = 1
        consecutive_early_count = 0
        while True:
            url = (f"{self.base_url}?pcform%5Bpublications_type%5D={pub_type}&pcform%5Bfield%5D=&pcform%5Bfrom_date_time%5D=&pcform%5Bto_date_time%5D=" if page == 1
                   else f"{self.base_url}?pcform%5Bfield%5D=&pcform%5Bfrom_date_time%5D=&pcform%5Bpublications_type%5D={pub_type}&pcform%5Bto_date_time%5D=&result-page={page}")
            print(f"\n  📄 第 {page} 頁")
            try:
                driver.get(url)
                time.sleep(2)
                articles = driver.find_elements(By.CSS_SELECTOR,
                    "div.content_container a[href*='/en/publications/cristin-pub/']")
                if not articles:
                    break
                for article_url in list(set([a.get_attribute('href') for a in articles])):
                    article_data, article_date = self.scrape_article(driver, article_url, start_dt, end_dt)
                    if article_data:
                        self.results.append(article_data)
                        consecutive_early_count = 0
                        print(f"  ✅ 收錄: {article_data['標題'][:50]}")
                    elif article_date and article_date < start_dt:
                        consecutive_early_count += 1
                        if consecutive_early_count >= 10:
                            print(f"\n  🛑 連續十篇早於起始時間，停止")
                            return
                    else:
                        consecutive_early_count = 0
                    time.sleep(1)
                page += 1
            except Exception as e:
                break

    def run(self, start_dt, end_dt):
        print("=" * 60)
        print("NUPI 爬蟲程式")
        print("=" * 60)
        driver = webdriver.Chrome(options=build_chrome_options())
        try:
            for pub_type in self.publication_types:
                self.scrape_publication_type(driver, pub_type, start_dt, end_dt)
        except Exception as e:
            print(f"\n❌ 執行時發生錯誤: {str(e)}")
        finally:
            driver.quit()
        print("\n🎉 NUPI 爬取完成！")


# ══════════════════════════════════════════════════════════════════════
# ORF Scraper
# ══════════════════════════════════════════════════════════════════════

class ORFScraper:
    def __init__(self):
        self.driver = None
        self.urls = [
            "https://www.orfonline.org/content-type/special-reports",
            "https://www.orfonline.org/content-type/books",
            "https://www.orfonline.org/content-type/commentary",
            "https://www.orfonline.org/content-type/occasional-paper",
            "https://www.orfonline.org/content-type/issue-briefs",
        ]
        self.keywords = ["China", "Taiwan", "Taipei"]
        self.articles_data = []

    def parse_date(self, date_str):
        try:
            return datetime.strptime(date_str.strip(), '%b %d, %Y')
        except Exception:
            return None

    def setup_driver(self):
        self.driver = webdriver.Chrome(options=build_chrome_options())
        self.driver.implicitly_wait(10)

    def check_keywords_in_text(self, text):
        if not text:
            return []
        text_upper = text.upper()
        return [kw for kw in self.keywords if kw.upper() in text_upper]

    def check_article(self, article_url):
        try:
            self.driver.get(article_url)
            time.sleep(3)
            title = ""
            for by, selector in [(By.CSS_SELECTOR, "h1"), (By.CSS_SELECTOR, "h3.heading_h3")]:
                try:
                    title = self.driver.find_element(by, selector).text.strip()
                    if title:
                        break
                except:
                    continue
            content = ""
            try:
                body = self.driver.find_element(By.TAG_NAME, "body")
                p_texts = [el.text for el in body.find_elements(By.TAG_NAME, "p") if el.text.strip()]
                content = " ".join(p_texts)
            except:
                pass
            found_keywords = self.check_keywords_in_text(title + " " + content)
            if not found_keywords:
                return None
            has_pdf = False
            for link in self.driver.find_elements(By.TAG_NAME, "a"):
                href = link.get_attribute("href")
                if href and href.lower().endswith('.pdf') and 'orfonline.org' in href:
                    has_pdf = True
                    break
            if has_pdf:
                return {'title': title, 'keywords': ', '.join(found_keywords)}
            return None
        except Exception:
            return None

    def scrape_page(self, url, start_dt, end_dt):
        print(f"{'-' * 20}")
        print(f"正在爬取: {url}")
        print(f"{'-' * 20}")
        consecutive_old = 0
        page_num = 1
        while consecutive_old < 3:
            page_url = url if page_num == 1 else f"{url}?page={page_num}"
            self.driver.get(page_url)
            time.sleep(3)
            article_blocks = self.driver.find_elements(By.CLASS_NAME, "all_story")
            if not article_blocks:
                break
            article_count = len(article_blocks)
            for i in range(article_count):
                try:
                    current_blocks = self.driver.find_elements(By.CLASS_NAME, "all_story")
                    if i >= len(current_blocks):
                        break
                    block_text = current_blocks[i].text
                    date_match = re.search(r'([A-Z][a-z]{2}\s+\d{1,2},\s+\d{4})', block_text)
                    if not date_match:
                        continue
                    article_date = self.parse_date(date_match.group(1))
                    if not article_date:
                        continue
                    if article_date < start_dt:
                        consecutive_old += 1
                        if consecutive_old >= 3:
                            return
                        continue
                    elif article_date > end_dt:
                        consecutive_old = 0
                        continue
                    else:
                        consecutive_old = 0
                    current_blocks = self.driver.find_elements(By.CLASS_NAME, "all_story")
                    if i >= len(current_blocks):
                        continue
                    try:
                        link_element = current_blocks[i].find_element(By.CSS_SELECTOR, "a[href*='/research/']")
                        article_url = link_element.get_attribute("href")
                        if not article_url or article_url == "#":
                            continue
                        article_info = self.check_article(article_url)
                        if article_info:
                            self.articles_data.append({
                                '機構': 'Observer Research Foundation',
                                '日期': article_date.strftime('%Y/%m/%d'),
                                '標題': article_info['title'],
                                '網址': article_url,
                                '關鍵字': article_info['keywords']
                            })
                            print(f"  ✓ 收錄: {article_info['title'][:50]}")
                        self.driver.get(page_url)
                        time.sleep(2)
                    except Exception:
                        try:
                            self.driver.get(page_url)
                            time.sleep(2)
                        except:
                            pass
                except Exception:
                    continue
            page_num += 1

    def run(self, start_dt, end_dt):
        print("=" * 60)
        print("ORF 網站爬蟲")
        print("=" * 60)
        self.setup_driver()
        try:
            for url in self.urls:
                self.scrape_page(url, start_dt, end_dt)
        except Exception as e:
            print(f"\n程式執行時發生錯誤: {e}")
        finally:
            if self.driver:
                self.driver.quit()
        print("\n🎉 ORF 爬取完成！")


# ══════════════════════════════════════════════════════════════════════
# IFANS Scraper
# ══════════════════════════════════════════════════════════════════════

class IFANSScraper:
    def __init__(self):
        self.institution = "Institute of Foreign Affairs and National Security"
        self.keywords = ["China", "Taiwan", "Taipei"]
        self.base_urls = {
            "IFANS Focus": "https://www.ifans.go.kr/knda/ifans/eng/pblct/PblctList.do?menuCl=P11&pageIndex=",
            "IFANS PERSPECTIVES": "https://www.ifans.go.kr/knda/ifans/eng/pblct/PblctList.do?menuCl=P19&pageIndex=",
            "IFANS STUDY REPORT": "https://www.ifans.go.kr/knda/ifans/eng/pblct/PblctList.do?menuCl=P20&pageIndex="
        }
        self.results = []

    def parse_date(self, date_str):
        try:
            date_str = date_str.replace('Upload Date', '').strip().strip('*')
            for fmt in ['%Y-%m-%d', '%d %B %Y', '%Y/%m/%d', '%d/%m/%Y', '%B %d, %Y']:
                try:
                    return datetime.strptime(date_str.strip(), fmt)
                except ValueError:
                    continue
            return None
        except:
            return None

    def check_keywords(self, text):
        if not text:
            return []
        text_upper = text.upper()
        return [kw for kw in self.keywords if kw.upper() in text_upper]

    def has_pdf_download(self, driver):
        try:
            time.sleep(1)
            download_links = driver.find_elements(By.XPATH,
                "//a[contains(@href, 'FileDownloadView') and contains(@href, '.pdf')]")
            if download_links:
                return True
            page_source = driver.page_source
            if '.pdf' in page_source and 'FileDownloadView' in page_source:
                return True
            return False
        except Exception:
            return False

    def extract_article_info_from_li(self, li_element):
        try:
            title_link = li_element.find_element(By.TAG_NAME, 'a')
            title = title_link.text.strip().split('\n')[0]
            onclick = title_link.get_attribute('onclick')
            if not onclick:
                return None
            match = re.search(r"fnCmdView\('(\d+)','([^']+)'\)", onclick)
            if not match:
                return None
            article_id = match.group(1)
            menu_cl = match.group(2)
            html = li_element.get_attribute('outerHTML')
            date_match = re.search(r'Upload Date[^0-9]*(\d{4}-\d{2}-\d{2})', html)
            if not date_match:
                return None
            return {'title': title, 'article_id': article_id, 'menu_cl': menu_cl, 'date_text': date_match.group(1)}
        except Exception:
            return None

    def build_article_url(self, article_id, menu_cl):
        return (f"https://www.ifans.go.kr/knda/ifans/eng/pblct/PblctView.do?"
                f"csrfPreventionSalt=null&pblctDtaSn={article_id}&menuCl={menu_cl}"
                f"&clCode={menu_cl}&koreanEngSe=ENG&pclCode=&chcodeId=&searchCondition=searchAll"
                f"&searchKeyword=&pageIndex=1")

    def scrape_article(self, driver, article_data, start_dt, end_dt, consecutive_old):
        try:
            article_url = self.build_article_url(article_data['article_id'], article_data['menu_cl'])
            title = article_data['title']
            print(f"\n  文章: {title[:60]}...")
            article_date = self.parse_date(article_data['date_text'])
            if not article_date:
                return consecutive_old
            print(f"    日期: {article_date.strftime('%Y/%m/%d')}")
            if article_date < start_dt:
                return consecutive_old + 1
            if article_date > end_dt:
                return 0
            driver.get(article_url)
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            except TimeoutException:
                return consecutive_old
            time.sleep(2)
            content_text = ""
            for selector in ["//div[contains(@class, 'editor')]", "//div[contains(@class, 'board_con')]",
                             "//article", "//main"]:
                try:
                    content_text = driver.find_element(By.XPATH, selector).text
                    if content_text and len(content_text) > 100:
                        break
                except:
                    continue
            found_keywords = self.check_keywords(title + " " + content_text)
            if not found_keywords:
                return consecutive_old
            if not self.has_pdf_download(driver):
                return consecutive_old
            self.results.append({
                '機構': self.institution,
                '日期': article_date.strftime('%Y/%m/%d'),
                '標題': title,
                '網址': article_url,
                '關鍵字': ', '.join(found_keywords)
            })
            print(f"    ✓✓✓ 收錄成功！")
            return 0
        except Exception as e:
            return consecutive_old

    def scrape_category(self, driver, category_name, base_url, start_dt, end_dt):
        print(f"\n{'-' * 20}")
        print(f"開始爬取: {category_name}")
        print(f"{'-' * 20}")
        page = 1
        consecutive_old_articles = 0
        while True:
            url = f"{base_url}{page}"
            print(f"\n第 {page} 頁...")
            try:
                driver.get(url)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                time.sleep(3)
                article_elements = driver.find_elements(By.XPATH, "//li[.//a[contains(@onclick, 'fnCmdView')]]")
                if not article_elements:
                    break
                articles_data = [self.extract_article_info_from_li(elem) for elem in article_elements]
                articles_data = [a for a in articles_data if a]
                for article_data in articles_data:
                    consecutive_old_articles = self.scrape_article(
                        driver, article_data, start_dt, end_dt, consecutive_old_articles)
                    if consecutive_old_articles >= 3:
                        print("連續3篇文章早於起始時間，停止此分類")
                        return
                    time.sleep(1)
                page += 1
            except Exception as e:
                break

    def run(self, start_dt, end_dt):
        print("=" * 60)
        print("IFANS 網站爬蟲")
        print("=" * 60)
        opts = build_chrome_options()
        opts.add_argument('--lang=en-US')
        if USE_WEBDRIVER_MANAGER:
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=opts)
        else:
            driver = webdriver.Chrome(options=opts)
        print("瀏覽器已啟動\n")
        try:
            for category_name, base_url in self.base_urls.items():
                self.scrape_category(driver, category_name, base_url, start_dt, end_dt)
        except Exception as e:
            print(f"\n程式執行時發生錯誤: {str(e)}")
        finally:
            driver.quit()
        print("\n🎉 IFANS 爬取完成！")


# ══════════════════════════════════════════════════════════════════════
# ISPI Scraper
# ══════════════════════════════════════════════════════════════════════

class ISPIScraper:
    BASE_URL  = "https://www.ispionline.it/en/publications"
    INSTITUTE = "Institute for International Political Studies"
    KEYWORDS  = ["China", "Taiwan", "Taipei"]
    CONSECUTIVE_OLD_LIMIT = 3
    MONTHS = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
               "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}

    def __init__(self):
        self.driver  = None
        self.results = []

    def _make_driver(self):
        self.driver = webdriver.Chrome(options=build_chrome_options())
        self.driver.implicitly_wait(6)

    def _wait_cards(self, timeout=15):
        WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.card-pubblicazione")))

    def _parse_article_date(self, text):
        m = re.search(r'(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})', text.strip())
        if m:
            day, mon, year = int(m.group(1)), m.group(2).lower(), int(m.group(3))
            if mon in self.MONTHS:
                try:
                    return datetime(year, self.MONTHS[mon], day)
                except ValueError:
                    pass
        return None

    def _get_cards(self):
        cards = []
        for item in self.driver.find_elements(By.CSS_SELECTOR, "div.card-pubblicazione"):
            art_type = ""
            try:
                art_type = item.find_element(By.CSS_SELECTOR, "a.btn-category").text.strip().upper()
            except:
                pass
            try:
                url = item.find_element(By.XPATH, ".//a[contains(@href,'/en/publication/')]").get_attribute("href")
            except:
                continue
            date_text = ""
            try:
                date_text = item.find_element(By.CSS_SELECTOR, "span.date").text.strip()
            except:
                pass
            cards.append((url, date_text, art_type))
        return cards

    def _scrape_article(self, url):
        self.driver.get(url)
        time.sleep(1.5)
        title = ""
        for sel in ["h1.title-55", "h1.seon-lookup-checked", "h1"]:
            try:
                title = self.driver.find_element(By.CSS_SELECTOR, sel).text.strip()
                if title:
                    break
            except:
                pass
        body = ""
        for sel in ["div.block-article-content", "article.container", "main#primary"]:
            try:
                body = self.driver.find_element(By.CSS_SELECTOR, sel).text
                if body:
                    break
            except:
                pass
        combined = title + "\n" + body
        found_kw = [kw for kw in self.KEYWORDS if re.search(rf'\b{kw}\b', combined, re.IGNORECASE)]
        has_dl = bool(self.driver.find_elements(By.XPATH,
            "//*[self::a or self::button][starts-with(translate(normalize-space(.),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'download')]"))
        return title, found_kw, has_dl

    def _wait_pagination(self, timeout=15):
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "ul.pagination, ul[data-tot-page]")))
            time.sleep(0.5)
        except:
            pass

    def run(self, start_dt, end_dt):
        print("=" * 60)
        print("Institute for International Political Studies文章爬蟲")
        print("=" * 60)
        self._make_driver()
        try:
            self.driver.get(self.BASE_URL)
            self._wait_cards()
            page_num = 1
            while True:
                print(f"  [頁面 {page_num}] 掃描文章列表…")
                cards = self._get_cards()
                if not cards:
                    break
                to_scrape  = []
                consecutive_old = 0
                stop_paging = False
                for url, date_text, art_type in cards:
                    if art_type in ("PODCAST", "VIDEO"):
                        continue
                    art_date = self._parse_article_date(date_text)
                    if art_date is None:
                        continue
                    if art_date > end_dt:
                        consecutive_old = 0
                        continue
                    if art_date < start_dt:
                        consecutive_old += 1
                        if consecutive_old >= self.CONSECUTIVE_OLD_LIMIT:
                            stop_paging = True
                            break
                        continue
                    consecutive_old = 0
                    to_scrape.append((url, art_date.strftime("%Y/%m/%d")))
                for url, date_str in to_scrape:
                    print(f"    ✓ 進入文章 ({date_str}): {url}")
                    title, found_kw, has_dl = self._scrape_article(url)
                    if found_kw and has_dl:
                        kw_str = ", ".join(found_kw)
                        print(f"      → 收錄！關鍵字: {kw_str}")
                        self.results.append({
                            '機構': self.INSTITUTE,
                            '日期': date_str,
                            '標題': title,
                            '網址': url,
                            '關鍵字': kw_str
                        })
                if stop_paging:
                    print("\n  連續三篇超出起始日期，停止。")
                    break
                print(f"  翻到第 {page_num + 1} 頁…")
                self.driver.get(self.BASE_URL)
                self._wait_cards()
                for target in range(2, page_num + 2):
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(1)
                    self._wait_pagination()
                    try:
                        li = self.driver.find_element(By.CSS_SELECTOR, f"li.js-btn-pag[data-paged='{target}']")
                        self.driver.execute_script("arguments[0].click();", li)
                        time.sleep(3)
                        self._wait_cards()
                    except Exception:
                        break
                else:
                    page_num += 1
                    continue
                break
        except Exception as e:
            print(f"\n❌ 執行時發生錯誤: {str(e)}")
        finally:
            if self.driver:
                self.driver.quit()
                print("  瀏覽器已關閉。")
        print("\n🎉 ISPI 爬取完成！")


# ══════════════════════════════════════════════════════════════════════
# 合併輸出 Excel
# ══════════════════════════════════════════════════════════════════════

def export_combined_excel(kiep_results, diis_articles, iai_results, nupi_results,
                           orf_results, ifans_results, ispi_results,
                           filename='combined_articles_G2_normal.xlsx'):
    def dedup(rows, url_key='網址', kw_key='關鍵字'):
        unique = {}
        for row in rows:
            url = row[url_key]
            if url in unique:
                existing = unique[url][kw_key]
                new_kw = row[kw_key]
                if new_kw not in existing:
                    unique[url][kw_key] = f"{existing}, {new_kw}"
            else:
                unique[url] = row.copy()
        return list(unique.values())

    # DIIS 格式轉換 + 去重
    diis_converted = [{
        '機構': a['institution'], '日期': a['date'],
        '標題': a['title'], '網址': a['url'], '關鍵字': a['keyword']
    } for a in diis_articles]

    all_rows = (kiep_results +
                dedup(diis_converted) +
                dedup(iai_results) +
                dedup(nupi_results) +
                dedup(orf_results) +
                dedup(ifans_results) +
                dedup(ispi_results))

    if not all_rows:
        print("⚠ 沒有找到任何符合條件的文章，不產生 Excel")
        return 0

    wb = Workbook()
    ws = wb.active
    ws.title = "爬取結果"
    headers = ['機構', '日期', '標題', '網址', '關鍵字']
    ws.append(headers)
    header_font = Font(bold=True, size=12)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for row in all_rows:
        ws.append([row['機構'], row['日期'], row['標題'], row['網址'], row['關鍵字']])
    ws.freeze_panes = 'A2'
    wb.save(filename)
    print(f"✅ Excel 已儲存: {filename}（共 {len(all_rows)} 筆）")
    return len(all_rows)


# ══════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════

def main():
    kiep_results  = []
    diis_articles = []
    iai_results   = []
    nupi_results  = []
    orf_results   = []
    ifans_results = []
    ispi_results  = []

    for name, cls, result_attr in [
        ("KIEP",  KIEPCrawler,  "results"),
        ("DIIS",  DIISScraper,  "articles"),
        ("IAI",   IAIScraper,   "results"),
        ("NUPI",  NUPIScraper,  "results"),
        ("ORF",   ORFScraper,   "articles_data"),
        ("IFANS", IFANSScraper, "results"),
        ("ISPI",  ISPIScraper,  "results"),
    ]:
        instance = cls()
        try:
            instance.run(start_date, end_date)
        except Exception as e:
            print(f"❌ {name} 爬取出錯: {e}")
        finally:
            try:
                if hasattr(instance, 'driver') and instance.driver:
                    instance.driver.quit()
            except:
                pass
        result = getattr(instance, result_attr, [])
        if name == "KIEP":   kiep_results  = result
        elif name == "DIIS": diis_articles = result
        elif name == "IAI":  iai_results   = result
        elif name == "NUPI": nupi_results  = result
        elif name == "ORF":  orf_results   = result
        elif name == "IFANS":ifans_results = result
        elif name == "ISPI": ispi_results  = result

    print("\n📊 正在匯出合併 Excel...")
    excel_path = "combined_articles_G2_normal.xlsx"
    total_count = export_combined_excel(
        kiep_results, diis_articles, iai_results, nupi_results,
        orf_results, ifans_results, ispi_results, excel_path)

    print("\n📧 正在寄送 Email...")
    log_text = logger.get_log()
    send_email(excel_path, log_text, total_count)

    print("🎉 全部完成！")


if __name__ == '__main__':
    main()
