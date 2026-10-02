# -*- coding: utf-8 -*-
import io
import sys
import os
import re
import time
import calendar
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

try:
    from webdriver_manager.chrome import ChromeDriverManager
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
    msg["Subject"] = f"G3爬蟲結果 {today.strftime('%Y/%m/%d')}（區間 {start_date.date()} ~ {end_date.date()}）"

    body = (
        f"G3爬蟲執行完畢。\n"
        f"執行日期：{today.strftime('%Y/%m/%d %H:%M')}\n"
        f"資料區間：{start_date.date()} ~ {end_date.date()}\n"
        f"共收錄：{total_count} 筆\n\n"
        f"詳細 log 請見附件 G3_log.txt，Excel 結果請見附件。"
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
    log_part.add_header("Content-Disposition", "attachment; filename=G3_log.txt")
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


def make_driver():
    opts = build_chrome_options()
    if USE_WEBDRIVER_MANAGER:
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=opts)
    return webdriver.Chrome(options=opts)


# ─────────────────────────────────────────────
# RIIR 爬蟲
# ─────────────────────────────────────────────

def _riir_parse_date(date_str):
    try:
        return datetime.strptime(date_str.strip(), '%d %B %Y')
    except Exception:
        return None


def _riir_get_article_info(driver, url):
    print(f"    正在檢查文章頁面...")
    for attempt in range(2):
        try:
            driver.get(url)
            time.sleep(5)
            break
        except Exception:
            if attempt == 1:
                return {'has_pdf': False, 'keywords': []}
            time.sleep(2)

    info = {'has_pdf': False, 'keywords': []}

    try:
        current_url = driver.current_url
        if current_url.lower().endswith('.pdf'):
            info['has_pdf'] = True
            try:
                driver.back()
                time.sleep(3)
                page_source_lower = driver.page_source.lower()
                keywords_found = []
                for kw in ['China', 'Taiwan', 'Taipei']:
                    if kw.lower() in page_source_lower:
                        keywords_found.append(kw)
                info['keywords'] = keywords_found
            except:
                pass
            return info

        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "article")))
            time.sleep(2)
        except:
            pass

        page_source = driver.page_source
        try:
            article_content = driver.find_element(By.TAG_NAME, 'article').text.lower()
        except:
            article_content = page_source.lower()

        keywords_found = [kw for kw in ['China', 'Taiwan', 'Taipei'] if kw.lower() in article_content]
        info['keywords'] = keywords_found
        if keywords_found:
            print(f"    找到關鍵字: {', '.join(keywords_found)}")

        page_source_lower = page_source.lower()
        pdf_keywords = ['view pdf', 'view the pdf', 'download pdf', 'download the pdf',
                        'pdf version', 'full text can be read', 'full text can be downloaded']
        for kw in pdf_keywords:
            if kw in page_source_lower:
                info['has_pdf'] = True
                print(f"    找到 PDF: {kw}")
                return info

        try:
            for link in driver.find_elements(By.TAG_NAME, 'a'):
                href = (link.get_attribute('href') or '').lower()
                text = (link.text or '').lower().strip()
                if '.pdf' in href and href.endswith('.pdf'):
                    info['has_pdf'] = True
                    return info
                if text and href:
                    if ('pdf' in text or 'full text' in text) and any(kw in text for kw in ['view', 'download', 'read']):
                        info['has_pdf'] = True
                        return info
        except:
            pass

    except Exception as e:
        print(f"    錯誤: {e}")

    return info


def scrape_riir(start_dt, end_dt):
    print("=" * 60)
    print("Royal Institute for International Relations文章爬蟲")
    print("=" * 60)

    driver = make_driver()
    articles = []
    page = 1
    consecutive_old = 0

    try:
        while True:
            url = ("https://www.egmontinstitute.be/publications/" if page == 1
                   else f"https://www.egmontinstitute.be/publications/page/{page}/")
            print(f"\n{'-' * 20}")
            print(f"正在檢查第 {page} 頁: {url}")
            print('-' * 20)
            driver.get(url)
            time.sleep(3)

            article_elements = driver.find_elements(By.CSS_SELECTOR, 'article.post-publication')
            if not article_elements:
                print("⚠ 沒有找到更多文章，停止爬取")
                break

            print(f"本頁找到 {len(article_elements)} 篇文章\n")
            articles_info = []

            for article in article_elements:
                try:
                    try:
                        date_elem = article.find_element(By.TAG_NAME, 'time')
                        article_date = _riir_parse_date(date_elem.text.strip())
                        if not article_date:
                            continue
                    except:
                        continue

                    try:
                        link_elem = article.find_element(By.CSS_SELECTOR, 'a.post-publication__title')
                        article_url = link_elem.get_attribute('href')
                        title = article.find_element(By.TAG_NAME, 'h3').text.strip()
                    except:
                        try:
                            for link in article.find_elements(By.TAG_NAME, 'a'):
                                href = link.get_attribute('href')
                                if href and not link.find_elements(By.TAG_NAME, 'img'):
                                    article_url = href
                                    title = article.find_element(By.TAG_NAME, 'h3').text.strip()
                                    break
                            else:
                                continue
                        except:
                            continue

                    articles_info.append({'title': title, 'url': article_url, 'date': article_date})
                except Exception:
                    continue

            for idx, article_data in enumerate(articles_info, 1):
                try:
                    article_date = article_data['date']
                    title = article_data['title']
                    article_url = article_data['url']

                    print(f"[{idx}/{len(articles_info)}] ", end='')

                    if article_date < start_dt:
                        consecutive_old += 1
                        print(f"日期 {article_date.strftime('%Y/%m/%d')} 早於起始時間 (連續第 {consecutive_old} 篇)")
                        if consecutive_old >= 3:
                            print(f"\n⚠ 連續 3 篇文章早於起始時間，停止爬取")
                            raise StopIteration
                        continue
                    elif article_date > end_dt:
                        print(f"日期 {article_date.strftime('%Y/%m/%d')} 晚於結束時間，跳過")
                        consecutive_old = 0
                        continue

                    consecutive_old = 0
                    print(f"檢查文章: {title[:50]}...")
                    print(f"    日期: {article_date.strftime('%Y/%m/%d')}")

                    info = _riir_get_article_info(driver, article_url)

                    if info['keywords'] and info['has_pdf']:
                        print(f"    ✓✓✓ 符合條件！")
                        articles.append({
                            '機構': 'Royal Institute for International Relations',
                            '日期': article_date.strftime('%Y/%m/%d'),
                            '標題': title,
                            '網址': article_url,
                            '關鍵字': ', '.join(info['keywords'])
                        })
                    else:
                        print(f"    ✗ 不符合條件")

                except StopIteration:
                    raise
                except Exception as e:
                    print(f"    處理錯誤: {e}")
                    continue

            page += 1

    except StopIteration:
        print("⚠ 爬取已停止")
    except Exception as e:
        print(f"\n發生錯誤: {e}")
    finally:
        try:
            driver.quit()
        except:
            pass

    print(f"RIIR 共找到 {len(articles)} 篇符合條件的文章")
    return articles


# ─────────────────────────────────────────────
# PIIE 爬蟲
# ─────────────────────────────────────────────

def _piie_parse_datetime(datetime_str):
    try:
        return datetime.strptime(datetime_str.replace('Z', '+00:00').split('+')[0], '%Y-%m-%dT%H:%M:%S')
    except:
        return None


def _piie_check_keywords(text):
    keywords = ['China', 'Taiwan', 'Taipei']
    return [kw for kw in keywords if kw.lower() in text.lower()]


def _piie_has_download_button(driver):
    try:
        elements = driver.find_elements(By.XPATH,
            "//a[@download] | //a[contains(translate(text(), 'DOWNLOAD', 'download'), 'download')]")
        return len(elements) > 0
    except:
        return False


def _piie_scrape_category(driver, category_url, start_dt, end_dt, results):
    print(f"\n{'-' * 20}")
    print(f"正在爬取: {category_url}")
    print(f"{'-' * 20}")

    driver.get(category_url)
    time.sleep(3)

    page_num = 1
    early_count = 0

    while True:
        print(f"\n[第 {page_num} 頁]")
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "article")))
        except TimeoutException:
            print("  ⚠ 找不到文章元素")
            break

        articles = driver.find_elements(By.TAG_NAME, "article")
        if not articles:
            break

        print(f"  找到 {len(articles)} 篇文章")

        for idx, article in enumerate(articles, 1):
            try:
                time_element = article.find_element(By.TAG_NAME, "time")
                datetime_str = time_element.get_attribute("datetime")
                if not datetime_str:
                    continue
                article_date = _piie_parse_datetime(datetime_str)
                if not article_date:
                    continue

                if article_date < start_dt:
                    early_count += 1
                    print(f"  [{idx}] ⏭ {article_date.strftime('%Y/%m/%d')} 早於起始日期（連續 {early_count}/3）")
                    if early_count >= 3:
                        print("\n  ⛔ 連續3篇早於起始日期，停止此分類")
                        return
                    continue
                elif article_date > end_dt:
                    early_count = 0
                    continue

                early_count = 0

                link_element = article.find_element(By.CSS_SELECTOR, "h2 a, h3 a, .teaser__title a")
                article_url = link_element.get_attribute("href")
                article_title = link_element.text.strip()

                print(f"  [{idx}] 📄 {article_title[:60]}...")

                driver.execute_script("window.open(arguments[0], '_blank');", article_url)
                driver.switch_to.window(driver.window_handles[-1])
                time.sleep(2)

                actual_url = driver.current_url
                if 'piie.com' not in actual_url:
                    print(f"       ⚠ 外部連結，跳過")
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                    continue

                page_text = driver.find_element(By.TAG_NAME, "body").text
                found_keywords = _piie_check_keywords(page_text)

                if found_keywords:
                    results.append({
                        '機構': 'Peterson Institute for International Economics',
                        '日期': article_date.strftime('%Y/%m/%d'),
                        '標題': article_title,
                        '網址': article_url,
                        '關鍵字': ', '.join(found_keywords)
                    })
                    print(f"       ✓ 關鍵字: {', '.join(found_keywords)}")
                else:
                    print(f"       ⚪ 未找到關鍵字")

                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(1)

            except Exception as e:
                print(f"  [{idx}] ❌ 錯誤: {str(e)[:50]}")
                if len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                continue

        try:
            next_buttons = driver.find_elements(By.XPATH,
                "//a[contains(translate(text(), 'NEXT', 'next'), 'next')] | //a[@rel='next']")
            if next_buttons:
                driver.execute_script("arguments[0].scrollIntoView();", next_buttons[0])
                time.sleep(1)
                next_buttons[0].click()
                time.sleep(3)
                page_num += 1
            else:
                break
        except Exception:
            break


def scrape_piie(start_dt, end_dt):
    print("=" * 60)
    print("Peterson Institute for International Economics文章爬蟲")
    print("=" * 60)

    categories = [
        'https://www.piie.com/research/commentary/testimonies',
        'https://www.piie.com/research/commentary/speeches-papers',
        'https://www.piie.com/publications/working-papers',
        'https://www.piie.com/publications/policy-briefs',
        'https://www.piie.com/publications/piie-briefings',
        'https://www.piie.com/bookstore'
    ]

    results = []
    driver = make_driver()

    try:
        for category_url in categories:
            try:
                _piie_scrape_category(driver, category_url, start_dt, end_dt, results)
            except Exception as e:
                print(f"\n❌ 爬取分類時發生錯誤: {e}")
                continue
    except Exception as e:
        print(f"\n❌ 執行過程中發生錯誤: {e}")
    finally:
        if driver:
            driver.quit()

    print(f"\nPIIE 共找到 {len(results)} 篇符合條件的文章")
    return results


# ─────────────────────────────────────────────
# CNAS 爬蟲
# ─────────────────────────────────────────────

def _cnas_parse_date(date_text):
    try:
        date_text = date_text.strip()
        date_text = ' '.join(date_text.split())
        for fmt in ['%B %d, %Y', '%b %d, %Y', '%Y-%m-%d', '%m/%d/%Y']:
            try:
                return datetime.strptime(date_text, fmt)
            except ValueError:
                continue
        return None
    except Exception:
        return None


def _cnas_check_keywords(driver):
    keywords = ['China', 'Taiwan', 'Taipei']
    found = []
    try:
        for selector in ['div[id*="biblio"]', '.more-from-cnas', 'footer', 'nav', 'aside']:
            try:
                driver.execute_script(f"""
                    document.querySelectorAll('{selector}').forEach(el => el.remove());
                """)
            except:
                pass
        content = driver.find_element(By.TAG_NAME, 'body').text
        for kw in keywords:
            if kw.lower() in content.lower():
                found.append(kw)
    except Exception:
        pass
    return found


def _cnas_check_pdf(driver):
    try:
        for selector in [
            "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download pdf')]",
            "//a[@class='button' and contains(., 'PDF')]",
        ]:
            elements = driver.find_elements(By.XPATH, selector)
            for elem in elements:
                text = elem.text.strip().lower()
                if 'pdf' in text and 'dataset' not in text:
                    return True
        return False
    except Exception:
        return False


def scrape_cnas(start_dt, end_dt):
    print("=" * 60)
    print("Center for a New American Security 文章爬蟲程式")
    print("=" * 60)

    categories = [
        'https://www.cnas.org/reports',
        'https://www.cnas.org/articles-multimedia?type=congressional-testimony',
        'https://www.cnas.org/articles-multimedia?type=commentary'
    ]

    results = []
    driver = make_driver()
    driver.implicitly_wait(10)

    try:
        for category_url in categories:
            print(f"\n{'-'*20}")
            print(f"正在爬取分類: {category_url}")
            print(f"{'-'*20}")

            page_num = 1
            consecutive_old = 0

            while True:
                current_url = (category_url if page_num == 1
                               else f"{category_url.split('?')[0]}/p{page_num}{'?' + category_url.split('?')[1] if '?' in category_url else ''}")
                print(f"\n📄 第 {page_num} 頁")
                driver.get(current_url)

                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'ul.entry-listing')))
                    time.sleep(2)
                except TimeoutException:
                    pass

                articles = driver.find_elements(By.CSS_SELECTOR, 'ul.entry-listing li')
                if not articles:
                    break

                for idx, article in enumerate(articles, 1):
                    try:
                        title_element = None
                        for selector in ['a.sans-serif', 'a[class*="sans-serif"]', 'a.fz16', 'div a', 'a']:
                            try:
                                title_element = article.find_element(By.CSS_SELECTOR, selector)
                                if title_element.text.strip():
                                    break
                            except NoSuchElementException:
                                continue

                        if not title_element or not title_element.text.strip():
                            continue

                        article_title = title_element.text.strip()
                        article_url = title_element.get_attribute('href')
                        if not article_url:
                            continue

                        article_date_text = None
                        article_text = article.text
                        date_pattern = r'(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER|January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}'
                        match = re.search(date_pattern, article_text)
                        if match:
                            article_date_text = match.group(0)

                        if not article_date_text:
                            continue

                        article_date = _cnas_parse_date(article_date_text)
                        if not article_date:
                            continue

                        if article_date < start_dt:
                            consecutive_old += 1
                            print(f"  ⏩ {article_title[:50]} - 早於起始時間（連續 {consecutive_old}/3）")
                            if consecutive_old >= 3:
                                print(f"\n✓ 連續 3 篇早於起始時間，停止此分類")
                                break
                            continue
                        elif article_date > end_dt:
                            consecutive_old = 0
                            continue

                        consecutive_old = 0
                        print(f"  📰 {article_title[:50]} | {article_date.strftime('%Y/%m/%d')}")

                        driver.get(article_url)
                        time.sleep(2)

                        found_keywords = _cnas_check_keywords(driver)
                        if not found_keywords:
                            print(f"     ❌ 未找到關鍵字")
                            driver.back()
                            time.sleep(1)
                            continue

                        has_pdf = _cnas_check_pdf(driver)
                        if not has_pdf:
                            print(f"     ❌ 無 PDF 下載按鈕")
                            driver.back()
                            time.sleep(1)
                            continue

                        results.append({
                            '機構': 'Center for a New American Security',
                            '日期': article_date.strftime('%Y/%m/%d'),
                            '標題': article_title,
                            '網址': article_url,
                            '關鍵字': ', '.join(found_keywords)
                        })
                        print(f"     ✅ 已收錄！關鍵字: {', '.join(found_keywords)}")
                        driver.back()
                        time.sleep(1)

                    except Exception:
                        continue

                if consecutive_old >= 3:
                    break

                page_num += 1
                if page_num > 10:
                    break

    except Exception as e:
        print(f"\n❌ 發生錯誤: {e}")
    finally:
        if driver:
            driver.quit()

    print(f"\nCNAS 共找到 {len(results)} 篇符合條件的文章")
    return results


# ─────────────────────────────────────────────
# EAI 爬蟲
# ─────────────────────────────────────────────

def _eai_parse_date(date_string):
    try:
        return datetime.strptime(date_string, '%Y-%m-%d')
    except:
        match = re.search(r'(\d{4}-\d{2}-\d{2})', date_string)
        if match:
            try:
                return datetime.strptime(match.group(1), '%Y-%m-%d')
            except:
                return None
        return None


def _eai_check_keywords(text):
    keywords = ['China', 'Taiwan', 'Taipei']
    found = []
    text_lower = text.lower()
    for kw in keywords:
        if re.search(r'\b' + re.escape(kw.lower()) + r'\b', text_lower):
            found.append(kw)
    return found


def _eai_scrape_article(driver, article_url, formatted_date, visited_urls):
    if article_url in visited_urls:
        return None
    visited_urls.add(article_url)

    try:
        driver.get(article_url)
        time.sleep(2)

        try:
            title = driver.find_element(By.CSS_SELECTOR, "div.tit").text.strip()
            print(f"      標題: {title[:60]}...")
        except:
            return None

        content = ""
        for selector in ["div.txt_wrap", "div.inner", "div.bt_wrap", "article", "main"]:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                for elem in elements:
                    content += elem.text.strip() + " "
            except:
                continue

        if len(content) < 500:
            try:
                content = driver.find_element(By.TAG_NAME, "body").text
            except:
                pass

        if "reference" in content.lower():
            content = content[:content.lower().find("reference")]

        found_keywords = _eai_check_keywords(title + " " + content)
        if not found_keywords:
            return None

        try:
            driver.find_element(By.CSS_SELECTOR, "a[href*='file_download.php']")
            has_pdf = True
        except NoSuchElementException:
            return None

        return {
            '機構': 'East Asia Institute',
            '日期': formatted_date,
            '標題': title,
            '網址': article_url,
            '關鍵字': ', '.join(found_keywords)
        }
    except Exception:
        return None


def _eai_scrape_listing(driver, keyword, start_dt, end_dt, visited_urls):
    articles = []
    start_param = 0
    consecutive_old = 0

    while True:
        url = f"https://eai.or.kr/eng/press/press_01.php?start={start_param}&category=&s_type=&s_keyword={keyword.lower()}"
        print(f"\n正在訪問: {url}")
        driver.get(url)
        time.sleep(3)

        try:
            article_rows = driver.find_elements(By.CSS_SELECTOR, "div.list_wrap div.row")
        except:
            break

        if not article_rows:
            break

        for idx in range(len(article_rows)):
            try:
                article_rows = driver.find_elements(By.CSS_SELECTOR, "div.list_wrap div.row")
                if idx >= len(article_rows):
                    continue
                row = article_rows[idx]

                try:
                    category = row.find_element(By.CSS_SELECTOR, "div.icon span").text.strip()
                except:
                    category = "Unknown"

                if category.lower() in ['multimedia', 'etc']:
                    continue

                try:
                    date_text = row.find_element(By.CSS_SELECTOR, "span.date").text.strip()
                except:
                    continue

                article_date = _eai_parse_date(date_text)
                if not article_date:
                    continue

                if article_date < start_dt:
                    consecutive_old += 1
                    if consecutive_old >= 5:
                        return articles
                    continue

                if article_date > end_dt:
                    continue

                consecutive_old = 0

                try:
                    article_url = row.find_element(By.CSS_SELECTOR, "a").get_attribute("href")
                except:
                    continue

                if not article_url.startswith("http"):
                    article_url = "https://eai.or.kr/eng/press/" + article_url

                result = _eai_scrape_article(driver, article_url, article_date.strftime('%Y/%m/%d'), visited_urls)
                if result:
                    articles.append(result)
                    print(f"    ✓ 已收錄文章！")

                driver.get(url)
                time.sleep(2)

            except Exception as e:
                driver.get(url)
                time.sleep(2)
                continue

        start_param += 8
        if start_param > 200:
            break

    return articles


def scrape_eai(start_dt, end_dt):
    print("=" * 60)
    print("East Asia Institute 文章爬蟲程序")
    print("=" * 60)

    driver = make_driver()
    driver.maximize_window()
    all_articles = []
    visited_urls = set()

    try:
        for keyword in ['china', 'taiwan', 'taipei']:
            print(f"\n{'-' * 20}")
            print(f"開始搜索關鍵字: {keyword.upper()}")
            print(f"{'-' * 20}")
            articles = _eai_scrape_listing(driver, keyword, start_dt, end_dt, visited_urls)
            all_articles.extend(articles)
            time.sleep(2)
    except Exception as e:
        print(f"\n發生錯誤: {e}")
    finally:
        driver.quit()

    print(f"\nEAI 共找到 {len(all_articles)} 篇符合條件的文章")
    return all_articles


# ─────────────────────────────────────────────
# LSE 爬蟲
# ─────────────────────────────────────────────

def scrape_lse(start_dt, end_dt):
    print("\n" + "=" * 60)
    print("LSE IDEAS 文章爬蟲程式")
    print("=" * 60)

    search_urls = {
        'China':  'https://researchonline.lse.ac.uk/cgi/tabbed_search/archive/simple?dataset=archive&screen=Search&exp=0%7C1%7C%7Carchive%7C-%7Cq%3A%3AALL%3AIN%3Achina%7C-%7C&order=-date%2Fcreators_name%2Ftitle',
        'Taiwan': 'https://researchonline.lse.ac.uk/cgi/tabbed_search/archive/simple?dataset=archive&screen=Search&exp=0%7C1%7C%7Carchive%7C-%7Cq%3A%3AALL%3AIN%3Ataiwan%7C-%7C&order=-date%2Fcreators_name%2Ftitle',
        'Taipei': 'https://researchonline.lse.ac.uk/cgi/tabbed_search/archive/simple?dataset=archive&screen=Search&exp=0%7C1%7C%7Carchive%7C-%7Cq%3A%3AALL%3AIN%3Ataipei%7C-%7C&order=-date%2Fcreators_name%2Ftitle'
    }

    opts = build_chrome_options()
    opts.page_load_strategy = 'eager'
    driver = webdriver.Chrome(options=opts) if not USE_WEBDRIVER_MANAGER else webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    driver.set_page_load_timeout(60)
    driver.implicitly_wait(10)

    results = []
    visited_urls = set()

    def safe_get(url, retries=3, wait=3):
        for attempt in range(retries):
            try:
                driver.get(url)
                time.sleep(wait)
                return True
            except Exception:
                time.sleep(5)
        return False

    def parse_date(date_str):
        try:
            return datetime.strptime(date_str, '%d %B %Y')
        except:
            return None

    try:
        for search_keyword, base_url in search_urls.items():
            offset = 0
            page_num = 1
            consecutive_old = 0

            print(f"\n{'-' * 20}")
            print(f"開始爬取關鍵字: {search_keyword}")
            print(f"{'-' * 20}")

            while True:
                url = base_url if offset == 0 else (
                    f"https://researchonline.lse.ac.uk/cgi/tabbed_search/archive/simple"
                    f"?exp=0%7C1%7C-date%2Fcreators_name%2Ftitle%7Carchive%7C-%7C"
                    f"q%3A%3AALL%3AIN%3A{search_keyword.lower()}%7C-%7C"
                    f"&_action_search=1&order=-date%2Fcreators_name%2Ftitle"
                    f"&screen=Material%3A%3ATabbedSearch&search_offset={offset}"
                )

                print(f"第 {page_num} 頁")
                if not safe_get(url, retries=3, wait=5):
                    break

                article_links = driver.find_elements(By.XPATH, "//a[contains(@href, '/id/eprint/')]")
                if not article_links:
                    break

                article_urls = [l.get_attribute('href') for l in article_links if l.get_attribute('href')]
                article_urls = [u for u in list(dict.fromkeys(article_urls)) if u and '/id/eprint/' in u]
                print(f"   找到 {len(article_urls)} 篇文章")

                stop_keyword = False
                for idx, article_url in enumerate(article_urls, 1):
                    if article_url in visited_urls:
                        continue
                    visited_urls.add(article_url)

                    if not safe_get(article_url, retries=3, wait=2):
                        continue

                    try:
                        title_elems = driver.find_elements(By.TAG_NAME, 'h1')
                        title = next((e.text.strip() for e in title_elems if e.text.strip() and len(e.text.strip()) > 5), "標題未找到")
                    except:
                        title = "標題未找到"

                    print(f"\n   [{idx}/{len(article_urls)}] {title}")

                    try:
                        date_elems = driver.find_elements(By.XPATH, "//th[contains(text(), 'Date Deposited')]/following-sibling::td")
                        if not date_elems:
                            continue
                        article_date = parse_date(date_elems[0].text.strip())
                        if not article_date:
                            continue
                    except:
                        continue

                    print(f"   📅 {article_date.strftime('%Y/%m/%d')}")

                    if article_date < start_dt:
                        consecutive_old += 1
                        if consecutive_old >= 5:
                            stop_keyword = True
                            break
                        continue
                    elif article_date > end_dt:
                        consecutive_old = 0
                        continue
                    else:
                        consecutive_old = 0

                    try:
                        page_text = driver.find_element(By.TAG_NAME, 'body').text
                        keywords_found = [kw for kw in ['China', 'Taiwan', 'Taipei'] if re.search(kw, page_text, re.IGNORECASE)]
                    except:
                        keywords_found = []

                    if not keywords_found:
                        continue

                    try:
                        has_pdf = bool(driver.find_elements(By.XPATH,
                            "//span[contains(@class, 'eprints-item--button-text') and "
                            "translate(normalize-space(text()), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ')='DOWNLOAD']"))
                    except:
                        has_pdf = False

                    if not has_pdf:
                        continue

                    results.append({
                        '機構': 'LSE IDEAS',
                        '日期': article_date.strftime('%Y/%m/%d'),
                        '標題': title,
                        '網址': article_url,
                        '關鍵字': ', '.join(keywords_found)
                    })
                    print(f"   ✅ 已加入結果（共 {len(results)} 篇）")

                if stop_keyword:
                    break

                offset += 20
                page_num += 1

    except Exception as e:
        print(f"\n執行時發生錯誤: {e}")
    finally:
        driver.quit()

    print(f"\nLSE 共找到 {len(results)} 篇符合條件的文章")
    return results


# ─────────────────────────────────────────────
# EPC 爬蟲
# ─────────────────────────────────────────────

def _epc_parse_date(date_string):
    date_string = date_string.strip()
    for fmt in ['%b %d, %Y', '%B %d, %Y', '%d %B %Y', '%d %b %Y', '%Y/%m/%d', '%Y-%m-%d']:
        try:
            return datetime.strptime(date_string, fmt)
        except ValueError:
            continue
    return None


def _epc_check_keywords(text):
    return [kw for kw in ['China', 'Taiwan', 'Taipei'] if kw.lower() in text.lower()]


def _epc_check_pdf_link(driver):
    pdf_patterns = [
        "Read the full version here", "Read the full Discussion Paper here",
        "Read the full paper here", "Read the full Policy Brief here",
        "Read the full publication here", "Download the full", "Read the full",
    ]
    try:
        for link in driver.find_elements(By.TAG_NAME, 'a'):
            try:
                href = link.get_attribute('href')
                if not href or '.pdf' not in href.lower():
                    continue
                link_text = link.text.strip()
                if 'here' in link_text.lower():
                    try:
                        parent_text = link.find_element(By.XPATH, '..').text
                        for pattern in pdf_patterns:
                            if pattern.lower() in parent_text.lower():
                                return True
                    except:
                        pass
            except Exception:
                continue
    except Exception:
        pass
    return False


def scrape_epc(start_dt, end_dt):
    print("\n" + "=" * 60)
    print("EPC 網站文章爬蟲")
    print("=" * 60)

    driver = make_driver()
    wait = WebDriverWait(driver, 10)
    articles = []
    visited_urls = set()

    try:
        for search_keyword in ['china', 'taiwan', 'taipei']:
            print(f"\n{'-' * 20}")
            print(f"正在搜索關鍵字: {search_keyword.upper()}")
            print(f"{'-' * 20}")
            page = 1
            consecutive_old = 0

            while True:
                url = f"https://www.epc.eu/search/?search_keywords={search_keyword}&tab=publications&p={page}"
                print(f"\n>>> 第 {page} 頁")
                driver.get(url)
                time.sleep(2)

                article_links = driver.find_elements(By.CSS_SELECTOR, 'a[href*="/publication/"]')
                if not article_links:
                    break

                article_urls = [l.get_attribute('href') for l in article_links
                                if l.get_attribute('href') and l.get_attribute('href') not in visited_urls]
                if not article_urls:
                    break

                for article_url in article_urls:
                    if article_url in visited_urls:
                        continue
                    visited_urls.add(article_url)

                    driver.get(article_url)
                    time.sleep(1)

                    try:
                        title = wait.until(EC.presence_of_element_located((By.TAG_NAME, 'h1'))).text.strip()
                        article_date = None

                        for sel_type, sel_val in [
                            (By.CSS_SELECTOR, '.publication-item-date'),
                            (By.CSS_SELECTOR, '.date'),
                            (By.XPATH, '//time'),
                        ]:
                            try:
                                date_text = driver.find_element(sel_type, sel_val).text.strip()
                                article_date = _epc_parse_date(date_text)
                                if article_date:
                                    break
                            except NoSuchElementException:
                                continue

                        if not article_date:
                            continue

                        if article_date < start_dt:
                            consecutive_old += 1
                            if consecutive_old >= 5:
                                raise StopIteration
                            continue
                        elif article_date > end_dt:
                            consecutive_old = 0
                            continue

                        consecutive_old = 0

                        content = driver.find_element(By.TAG_NAME, 'body').text
                        found_keywords = _epc_check_keywords(title + ' ' + content)
                        if not found_keywords:
                            continue

                        if not _epc_check_pdf_link(driver):
                            continue

                        articles.append({
                            '機構': 'European Policy Centre',
                            '日期': article_date.strftime('%Y/%m/%d'),
                            '標題': title,
                            '網址': article_url,
                            '關鍵字': ', '.join(found_keywords)
                        })
                        print(f"  ✓✓✓ 收錄: {title[:50]}")

                    except StopIteration:
                        raise
                    except Exception:
                        continue

                    driver.back()
                    time.sleep(1)

                page += 1

    except StopIteration:
        pass
    except Exception as e:
        print(f"爬取時發生錯誤: {e}")
    finally:
        driver.quit()

    print(f"\nEPC 共找到 {len(articles)} 篇符合條件的文章")
    return articles

# ─────────────────────────────────────────────
# 主程式
# ─────────────────────────────────────────────

def save_to_excel(all_results, output_path='combined_articles_G3_normal.xlsx'):
    if not all_results:
        print("\n未找到任何符合條件的文章，不產生 Excel 檔案")
        return 0

    df = pd.DataFrame(all_results)
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"\n✓ 共收錄 {len(all_results)} 篇文章，結果已儲存至: {output_path}")
    return len(all_results)


def main():
    all_results = []

    for name, func in [
        ("RIIR",  scrape_riir),
        ("PIIE",  scrape_piie),
        ("CNAS",  scrape_cnas),
        ("EAI",   scrape_eai),
        ("LSE",   scrape_lse),
        ("EPC",   scrape_epc),
    ]:
        try:
            results = func(start_date, end_date)
            all_results.extend(results)
        except Exception as e:
            print(f"\n❌ {name} 爬取失敗，跳過：{e}")

    excel_path = 'combined_articles_G3_normal.xlsx'
    total_count = save_to_excel(all_results, excel_path)

    print("\n📧 正在寄送 Email...")
    log_text = logger.get_log()
    send_email(excel_path, log_text, total_count)

    print("🎉 全部完成！")


if __name__ == "__main__":
    main()
