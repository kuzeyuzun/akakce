############################################################
# DOSYA BİLGİLERİ
############################################################
# Dosya Adı: akakce_scraper.py
# Dosya Yolu: C:/K/PY/akakce_scraper.py
# Python Yolu: C:/Users/kuzey.uzun/AppData/Local/Programs/Python/Python313/python.exe
############################################################

############################################################
# PROJE BİLGİLERİ
############################################################
# Proje Adı: Akakce TV Scraper
# Sürüm: 2.40
# Son Güncelleme: 2025-03-06 15:30:00 UTC
# Geliştirici: Kuzey Uzun
# Mevcut Kullanıcı: seydauzun
############################################################

#######################
# MODÜL AÇIKLAMALARI #
#######################
# requests: HTTP istekleri için
# bs4: HTML parsing için
# selenium: Web otomasyon için
# gspread: Google Sheets API için
# asyncio: Asenkron işlemler için
# concurrent.futures: Paralel işlem için
# logging: Log kaydı için
#######################

import requests
from typing import List, Dict, Any, Optional, Set
from dataclasses import dataclass
import time
import random
import asyncio
import os
import logging
import sys
from concurrent.futures import ThreadPoolExecutor
from functools import lru_cache
from datetime import datetime
from pathlib import Path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException

import gspread
from oauth2client.service_account import ServiceAccountCredentials

#######################
# URL YAPILANDIRMASI #
#######################
# Ana URL yapılandırması
DOMAIN = "https://www.akakce.com"
CATEGORY = "televizyon"
BRAND = "LG"  # Boş bırakılırsa tüm markalar listelenir

# Base URL'i oluştur
BASE_URL = f"{DOMAIN}/{CATEGORY}/{BRAND}" if BRAND else f"{DOMAIN}/{CATEGORY}"

#######################
# PROGRAM AYARLARI   #
#######################
# Kazıma Ayarları
MAX_PAGES_TO_SCRAPE = 1     # Kazınacak sayfa sayısı
WAIT_TIME = 30              # Sayfa yükleme bekleme süresi (saniye)
MAX_WORKERS = 3             # Paralel çalışacak tarayıcı sayısı
BATCH_SIZE = 50             # Toplu veri gönderme boyutu

# Google Sheets Ayarları
SPREADSHEET_ID = "1tMv7ElvqIBqmO6mqPCk_w-HprA9O02J4gVBGdXQ-4ns"
MAIN_SHEET_NAME = "Asorti_TV"

# Dosya ve Script Ayarları
PTT_SCRIPT_NAME = "Ptt_ID_Bul.py"
CREDENTIALS_FILE = "credentials.json"

# Chrome User Agent
CHROME_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/91.0.4472.124 Safari/537.36"
)

#######################
# GELİŞTİRİCİ AYARLARI #
#######################
RETRY_DELAY = 2     # Hata durumunda bekleme süresi
MAX_RETRIES = 3     # Maksimum deneme sayısı
LOG_LEVEL = 3       # Chrome log seviyesi (3=sadece hatalar)
DEBUG_MODE = True   # Debug modu açık/kapalı

# Logging yapılandırması
def setup_logging():
    """Logging sistemini yapılandırır"""
    try:
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        current_date = datetime.now().strftime("%Y%m%d")
        log_file = log_dir / f"akakce_scraper_{current_date}.log"
        
        logging.basicConfig(
            level=logging.DEBUG if DEBUG_MODE else logging.INFO,
            format='[%(asctime)s] %(levelname)s: %(message)s',
            datefmt='%H:%M:%S',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        return logging.getLogger(__name__)
    except Exception as e:
        print(f"Logging kurulum hatası: {str(e)}")
        sys.exit(1)

# Logger'ı başlat
logger = setup_logging()

def print_status(message: str, level: str = 'info'):
    """Gelişmiş durum mesajı yazdırma fonksiyonu"""
    if level == 'debug':
        logger.debug(message)
    elif level == 'warning':
        logger.warning(message)
    elif level == 'error':
        logger.error(message)
    else:
        logger.info(message)

@dataclass
class ProductData:
    """Ürün verisi için veri sınıfı"""
    name: str
    link: str
    akakce_sku: str
    sellers: List[Dict[str, str]]

class SafeWebDriver:
    """Thread-safe WebDriver yönetimi için context manager"""
    def __init__(self, options: Options):
        self.options = options
        self.driver = None

    def __enter__(self):
        try:
            self.driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=self.options
            )
            # Anti-bot algılama önlemi
            self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                'source': '''
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    })
                '''
            })
            return self.driver
        except Exception as e:
            print_status(f"Driver başlatma hatası: {str(e)}", 'error')
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            raise

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.driver:
            try:
                self.driver.quit()
                self.driver = None
            except Exception as e:
                print_status(f"Driver kapatma hatası: {str(e)}", 'error')

class AkakceScraper:
    def __init__(self, wait_time: int = WAIT_TIME):
        self.wait_time = wait_time
        self.chrome_options = self._setup_chrome_options()
        self.browser_pool = []
        self.seen_products: Set[str] = set()
        self.base_url = BASE_URL

    def _get_page_url(self, page_number: int) -> str:
        """Sayfa numarasına göre URL oluşturur"""
        if page_number > 1:
            return f"{self.base_url},{page_number}.html"
        return f"{self.base_url}.html"

    @lru_cache(maxsize=128)
    def _setup_chrome_options(self) -> Options:
        """Chrome seçeneklerini önbellekle"""
        options = Options()
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument(f"--log-level={LOG_LEVEL}")
        options.add_argument("--window-size=1920,1080")
        options.page_load_strategy = 'eager'
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument(f"user-agent={CHROME_USER_AGENT}")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        return options

    async def scrape_multiple_pages(self) -> List[Dict[str, Any]]:
        """Asenkron sayfa kazıma"""
        all_products = []
        
        try:
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futures = []
                for page_num in range(1, MAX_PAGES_TO_SCRAPE + 1):
                    futures.append(
                        executor.submit(self.scrape_akakce_page, page_num)
                    )
                
                for future in futures:
                    products = future.result()
                    if products:
                        all_products.extend(products)
                
                return all_products
                
        except Exception as e:
            print_status(f"Çoklu sayfa kazıma hatası: {str(e)}", 'error')
            return all_products

    def scrape_akakce_page(self, page_number: int) -> List[Dict[str, Any]]:
        """SafeWebDriver kullanarak optimize edilmiş sayfa kazıma"""
        products = []
        url = self._get_page_url(page_number)
        
        for attempt in range(MAX_RETRIES):
            try:
                with SafeWebDriver(self.chrome_options) as driver:
                    print_status(f"Sayfa yükleniyor: {url}")
                    driver.get(url)
                    
                    # Sayfa yüklenene kadar bekle
                    WebDriverWait(driver, self.wait_time).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'li.w'))
                    )
                    
                    # Sayfanın tam olarak yüklenmesi için kısa bir bekleme
                    time.sleep(2)
                    
                    # JavaScript ile scroll
                    driver.execute_script(
                        "window.scrollTo(0, document.body.scrollHeight);"
                        "window.scrollTo(0, 0);"
                    )
                    
                    # Sayfa kaynağını al ve parse et
                    soup = BeautifulSoup(driver.page_source, 'html.parser')
                    product_elements = soup.find_all('li', class_='w')
                    
                    if not product_elements:
                        # Ürün bulunamadıysa tekrar dene
                        print_status(f"Sayfa {page_number}: Ürün bulunamadı, yeniden deneniyor...", 'warning')
                        if attempt < MAX_RETRIES - 1:
                            time.sleep(RETRY_DELAY)
                            continue
                    
                    for element in product_elements:
                        try:
                            product_details = self._extract_product_details(element)
                            if product_details:
                                products.append(product_details)
                        except Exception as e:
                            print_status(f"Ürün çıkarma hatası: {str(e)}", 'debug')
                            continue

                    print_status(f"Sayfa {page_number}: {len(products)} ürün kazındı")
                    break  # Başarılı, döngüden çık
                
            except TimeoutException:
                print_status(f"Sayfa {page_number}: Zaman aşımı, yeniden deneniyor... ({attempt+1}/{MAX_RETRIES})", 'warning')
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY)
            except Exception as e:
                print_status(f"Sayfa {page_number} kazıma hatası: {str(e)}", 'error')
                if attempt < MAX_RETRIES - 1:
                    time.sleep(RETRY_DELAY)
                else:
                    break

        return products

    def _extract_product_details(self, product_element) -> Optional[Dict[str, Any]]:
        """Ürün detayı çıkarma, geliştirilmiş hata işleme ile"""
        try:
            product_name_tag = product_element.select_one("h3.pn_v8")
            if not product_name_tag:
                return None

            product_name = product_name_tag.get_text(strip=True)
            link_tag = product_name_tag.find_parent('a')
            
            if not link_tag or 'href' not in link_tag.attrs:
                print_status(f"'{product_name}' için link bulunamadı", 'debug')
                return None
                
            product_link = link_tag['href']

            # SKU çıkarma ve doğrulama
            try:
                akakce_sku = product_link.split(',')[-1].split('.html')[0]
                # SKU kontrolü ekle
                if not akakce_sku or not akakce_sku.strip() or akakce_sku in self.seen_products:
                    return None
            except IndexError:
                print_status(f"'{product_name}' için SKU çıkarma hatası: {product_link}", 'debug')
                return None

            self.seen_products.add(akakce_sku)

            # Tam URL oluştur
            if not product_link.startswith("http"):
                product_link = f"{DOMAIN}{product_link}"

            # Satıcı bilgilerini çıkar
            seller_details = []
            seller_elements = product_element.select("div.p_w_v9 a.iC")
            
            for seller in seller_elements[:3]:  # İlk 3 satıcıyı al
                seller_info = self._extract_seller_details(seller)
                if seller_info:
                    seller_details.append(seller_info)

            if not seller_details:
                print_status(f"'{product_name}' için satıcı bulunamadı", 'debug')
                return None

            return {
                "name": product_name,
                "link": product_link,
                "akakce_sku": akakce_sku,
                "sellers": seller_details
            }

        except Exception as e:
            print_status(f"Ürün detayları çıkarma hatası: {str(e)}", 'debug')
            return None

    def _extract_seller_details(self, seller_element) -> Optional[Dict[str, str]]:
        """Geliştirilmiş satıcı detayı çıkarma"""
        try:
            # Satıcı adı bulma
            seller_name = "Bilinmeyen"
            seller_name_tag = seller_element.select_one("span.l i img")
            
            if seller_name_tag and 'alt' in seller_name_tag.attrs:
                seller_name = seller_name_tag.get('alt')
            else:
                # Alternatif yöntem
                seller_name_span = seller_element.select_one("span.l")
                if seller_name_span:
                    seller_name = seller_name_span.get_text(strip=True)
            
            # Fiyat bulma
            price_tag = seller_element.select_one("span.pt_v8")
            if not price_tag:
                return None
            
            price = price_tag.get_text(strip=True)
            
            # Veri doğrulama
            if not price or not seller_name:
                return None
                
            return {"name": seller_name, "price": price}
            
        except Exception as e:
            print_status(f"Satıcı detayları çıkarma hatası: {str(e)}", 'debug')
            return None

class GoogleSheetsManager:
    def __init__(self):
        self.scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        self.client = None
        self.spreadsheet = None
        self.connected = False

    async def connect(self):
        """Google Sheets bağlantısını kur"""
        try:
            if os.path.exists(CREDENTIALS_FILE):
                creds = ServiceAccountCredentials.from_json_keyfile_name(
                    CREDENTIALS_FILE, self.scope
                )
                self.client = gspread.authorize(creds)
                self.spreadsheet = self.client.open_by_key(SPREADSHEET_ID)
                self.connected = True
                print_status("Google Sheets bağlantısı kuruldu")
            else:
                print_status(f"Kimlik dosyası bulunamadı: {CREDENTIALS_FILE}", 'error')
                raise FileNotFoundError(f"Kimlik dosyası bulunamadı: {CREDENTIALS_FILE}")
                
        except Exception as e:
            print_status(f"Google Sheets bağlantı hatası: {str(e)}", 'error')
            raise

    async def update_sheets(self, data: List[Dict[str, Any]]):
        """Asenkron sheet güncelleme"""
        if not self.connected:
            await self.connect()
            
        if not data:
            print_status("Güncellenecek veri bulunamadı", 'warning')
            return
            
        try:
            # Ana sayfayı al
            main_sheet = self.spreadsheet.worksheet(MAIN_SHEET_NAME)
            main_sheet.clear()

            # Başlıkları bir kere yaz
            header = [
                "Sıra No", "Ürün Adı", "Link",
                "Satıcı 1", "Fiyat 1", "Satıcı 2", "Fiyat 2",
                "Satıcı 3", "Fiyat 3", "Akakce SKU",
                "PttAVM ID", "PttAVM Link", "Mağaza Adı", "Satış Fiyatı"
            ]
            main_sheet.append_row(header)
            
            # Verileri hazırla
            all_values = []
            for idx, product in enumerate(data, 1):
                row = [idx, product["name"], product["link"]]
                
                # Satıcı bilgilerini ekle
                for i in range(3):
                    if i < len(product["sellers"]):
                        seller = product["sellers"][i]
                        row.extend([seller["name"], seller["price"]])
                    else:
                        row.extend(["", ""])
                
                # Akakce SKU'yu ekle
                row.append(product.get("akakce_sku", ""))
                
                # Boş sütunları ekle (PTT verisi için)
                row.extend(["", "", "", ""])  
                
                all_values.append(row)
                
                # Toplu güncelleme için kontrol
                if len(all_values) >= BATCH_SIZE:
                    try:
                        main_sheet.append_rows(all_values)
                        all_values = []
                    except Exception as e:
                        print_status(f"Batch güncelleme hatası: {str(e)}", 'error')
                        # Hata durumunda tek tek güncellemeyi dene
                        for single_row in all_values:
                            try:
                                main_sheet.append_row(single_row)
                                time.sleep(0.5)  # API limitlerine takılmamak için
                            except:
                                pass
                        all_values = []
            
            # Kalan verileri güncelle
            if all_values:
                main_sheet.append_rows(all_values)

            print_status(f"Google Sheets güncellendi. Toplam {len(data)} ürün.")

        except Exception as e:
            print_status(f"Google Sheets güncellenirken hata: {str(e)}", 'error')
            raise

async def check_ptt_script():
    """PTT scriptinin var olup olmadığını kontrol et"""
    if not os.path.exists(PTT_SCRIPT_NAME):
        print_status(f"UYARI: PTT script dosyası ({PTT_SCRIPT_NAME}) bulunamadı.", 'warning')
        return False
    return True

async def run_ptt_script():
    """PTT ID bulma scriptini çalıştır"""
    try:
        if not await check_ptt_script():
            return False
            
        print_status("PTT ID bulma işlemi başlatılıyor...")
        
        import subprocess
        process = subprocess.run(
            ["python", PTT_SCRIPT_NAME], 
            capture_output=True, 
            text=True,
            check=False
        )
        
        if process.returncode == 0:
            print_status("PTT ID bulma işlemi başarıyla tamamlandı.")
            return True
        else:
            print_status(f"PTT ID bulma işlemi başarısız oldu. Hata: {process.stderr}", 'error')
            return False
            
    except Exception as e:
        print_status(f"PTT script çalıştırma hatası: {str(e)}", 'error')
        return False

async def main():
    """Asenkron ana fonksiyon - geliştirilmiş hata yönetimi"""
    start_time = time.time()
    
    try:
        print_status("Akakce veri kazıma başlıyor...")
        scraper = AkakceScraper()
        products = await scraper.scrape_multiple_pages()
        
        if products:
            print_status(f"Toplam {len(products)} ürün kazındı.")
            print_status("Google Sheets güncelleniyor...")
            
            sheets_manager = GoogleSheetsManager()
            await sheets_manager.connect()
            await sheets_manager.update_sheets(products)
            
            # Akakce işlemi bittikten sonra PTT ID bulma işlemini başlat
            ptt_success = await run_ptt_script()
            
            end_time = time.time()
            duration = end_time - start_time
            print_status(f"Tüm işlemler tamamlandı. Toplam süre: {duration:.2f} saniye")
            
        else:
            print_status("Hiç ürün bulunamadı!", 'warning')

    except KeyboardInterrupt:
        print_status("Kullanıcı tarafından durduruldu.", 'warning')
    except Exception as e:
        print_status(f"Program hatası: {str(e)}", 'error')
        # Hata detaylarını logla ama yeniden fırlatma
        import traceback
        print_status(traceback.format_exc(), 'error')
    finally:
        end_time = time.time()
        duration = end_time - start_time
        print_status(f"Program sonlandı. Çalışma süresi: {duration:.2f} saniye")

if __name__ == "__main__":
    asyncio.run(main())