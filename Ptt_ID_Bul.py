############################################################
# DOSYA BİLGİLERİ
############################################################
# Dosya Adı: Ptt_ID_Bul.py
# Dosya Yolu: C:/K/PY/Ptt_ID_Bul.py
# Python Yolu: C:/Users/kuzey.uzun/AppData/Local/Programs/Python/Python313/python.exe
############################################################

############################################################
# PROJE BİLGİLERİ
############################################################
# Proje Adı: PTT AVM ID Bulucu
# Sürüm: 1.8
# Son Güncelleme: 2025-03-06 16:45:00 UTC
# Geliştirici: Kuzey Uzun
# Geliştirici Desteği: GitHub Copilot & Claude AI
# Mevcut Kullanıcı: seydauzun
############################################################

import os
import time
import logging
import traceback
from datetime import datetime, UTC
from pathlib import Path
from typing import Optional, List, Dict, Any
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor, as_completed

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Sabit Değişkenler
SPREADSHEET_ID = "1tMv7ElvqIBqmO6mqPCk_w-HprA9O02J4gVBGdXQ-4ns"
WORKSHEET_NAME = "Asorti_TV"
CREDENTIALS_FILE = "credentials.json"
WAIT_TIME = 10
MAX_RETRIES = 3
MAX_WORKERS = 5  # Paralel işlem için worker sayısı
DEBUG_MODE = True
CHROME_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

@dataclass
class ProductData:
    """Ürün verisi için veri sınıfı"""
    ptt_link: str = ""
    status: bool = False
    error: str = ""

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
            logger.error(f"Driver başlatma hatası: {str(e)}")
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
                logger.error(f"Driver kapatma hatası: {str(e)}")

def setup_logging():
    """Logging sistemini yapılandırır"""
    try:
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        
        current_date = datetime.now().strftime("%Y%m%d")
        log_file = log_dir / f"ptt_finder_{current_date}.log"
        
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
        raise

# Global logger
logger = setup_logging()

class PttLinkFetcher:
    def __init__(self):
        self.debug_mode = DEBUG_MODE
        
        # Google Sheets API yapılandırması
        self.scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        self.credentials_path = CREDENTIALS_FILE
        self.sheet_id = SPREADSHEET_ID
        self.worksheet_name = WORKSHEET_NAME
        
        # Chrome seçenekleri
        self.chrome_options = self._setup_chrome_options()
        
        # İstatistikler
        self.success_count = 0
        self.error_count = 0
        self.total_count = 0
        self.start_time = time.time()
        
    def _setup_chrome_options(self) -> Options:
        """Chrome seçeneklerini yapılandırır"""
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument("--window-size=1920,1080")
        options.add_argument(f"user-agent={CHROME_USER_AGENT}")
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.page_load_strategy = 'eager'  # Sayfanın tam yüklenmesini beklemeden devam et
        return options

    def connect_to_sheets(self):
        """Google Sheets bağlantısını kurar"""
        try:
            credentials = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_path, self.scope)
            gc = gspread.authorize(credentials)
            self.sheet = gc.open_by_key(self.sheet_id).worksheet(self.worksheet_name)
            logger.info("Google Sheets bağlantısı kuruldu")
            
            # İşlem başladı mesajı
            self.sheet.update_cell(1, 15, "Lütfen Bekleyiniz...")
            
        except Exception as e:
            logger.error(f"Sheets bağlantı hatası: {str(e)}")
            raise

    def fetch_ptt_link(self, url: str, retry_count: int = 0) -> ProductData:
        """Sayfadaki PTT AVM linkini bulur - geliştirilmiş hata yönetimi"""
        if not url or not url.startswith("http"):
            return ProductData(status=False, error="Geçersiz URL")
            
        if retry_count >= MAX_RETRIES:
            return ProductData(status=False, error="Maksimum deneme sayısı aşıldı")
        
        try:
            with SafeWebDriver(self.chrome_options) as driver:
                driver.get(url)
                wait = WebDriverWait(driver, WAIT_TIME)
                
                # Satıcı tablosunu bekle
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "sellers_table")))
                
                # Tüm linkleri kontrol et
                links = driver.find_elements(By.TAG_NAME, "a")
                for link in links:
                    try:
                        href = link.get_attribute('href')
                        if href and 'pttavm.com' in href:
                            return ProductData(ptt_link=href, status=True)
                    except:
                        continue
                
                # Alternative: Satıcı hücrelerini kontrol et
                seller_cells = driver.find_elements(By.CSS_SELECTOR, "td.sl_v2")
                for cell in seller_cells:
                    try:
                        link = cell.find_element(By.TAG_NAME, "a")
                        href = link.get_attribute('href')
                        if href and 'pttavm.com' in href:
                            return ProductData(ptt_link=href, status=True)
                    except:
                        continue
                
                # Link bulunamadı
                return ProductData(status=False, error="PTT link bulunamadı")
                
        except TimeoutException:
            # Zaman aşımı durumunda yeniden dene
            logger.warning(f"URL {url} için zaman aşımı, yeniden deneniyor ({retry_count+1}/{MAX_RETRIES})")
            time.sleep(2)
            return self.fetch_ptt_link(url, retry_count + 1)
            
        except Exception as e:
            logger.error(f"URL {url} için hata: {str(e)}")
            return ProductData(status=False, error=str(e))

    def process_row(self, row_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Tek bir satırı işler ve sonuç döndürür"""
        row = row_data["row"]
        url = row_data["url"]
        
        try:
            if not url:
                return None
                
            logger.info(f"İşleniyor: Satır {row} - {url}")
            
            result = self.fetch_ptt_link(url)
            
            if result.status and result.ptt_link:
                return {
                    "row": row,
                    "ptt_link": result.ptt_link,
                    "status": True
                }
            else:
                logger.warning(f"Satır {row}: {result.error}")
                return {
                    "row": row,
                    "ptt_link": "",
                    "status": False,
                    "error": result.error
                }
                
        except Exception as e:
            logger.error(f"Satır {row} işleme hatası: {str(e)}")
            return {
                "row": row,
                "ptt_link": "",
                "status": False,
                "error": str(e)
            }

    def update_sheet(self):
        """PTT linklerini paralel olarak günceller"""
        try:
            # Tüm verileri al
            all_values = self.sheet.get_all_records()
            
            if not all_values:
                logger.warning("Güncellenecek veri bulunamadı")
                return
                
            logger.info(f"Toplam {len(all_values)} satır işlenecek")
            
            # İşlenecek satırları hazırla
            rows_to_process = []
            for idx, record in enumerate(all_values, 2):  # 2'den başla çünkü Google Sheets'te başlık satırı 1
                url = record.get("Link", "")
                if url and url.startswith("http"):
                    rows_to_process.append({
                        "row": idx,
                        "url": url
                    })
            
            results = []
            
            # Paralel işleme başla
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                future_to_row = {executor.submit(self.process_row, row_data): row_data for row_data in rows_to_process}
                
                for future in as_completed(future_to_row):
                    result = future.result()
                    if result:
                        results.append(result)
                        
                        # Sonucu hemen güncelle
                        if result["status"]:
                            self.sheet.update_cell(result["row"], 12, result["ptt_link"])  # L sütunu
                            self.success_count += 1
                        else:
                            self.sheet.update_cell(result["row"], 12, "")  # L sütunu
                            self.error_count += 1
                            
                        self.total_count += 1
                        
                        # Süresini güncelle
                        current_time = time.time()
                        elapsed = current_time - self.start_time
                        avg_time = elapsed / self.total_count if self.total_count > 0 else 0
                        remaining = avg_time * (len(rows_to_process) - self.total_count)
                        
                        status_msg = f"İşlenen: {self.total_count}/{len(rows_to_process)} - Kalan süre: {remaining:.1f}s"
                        self.sheet.update_cell(1, 15, status_msg)
                        
            logger.info(f"Tüm satırlar işlendi. Başarılı: {self.success_count}, Hata: {self.error_count}")
            
        except Exception as e:
            logger.error(f"Genel güncelleme hatası: {str(e)}")
            logger.error(traceback.format_exc())
            raise
        finally:
            try:
                self.sheet.update_cell(1, 15, "")  # İşlem tamamlandı mesajını temizle
            except:
                pass

def main():
    start_time = time.time()
    try:
        logger.info("Program başlatılıyor...")
        logger.info(f"Başlangıç zamanı: {datetime.now(UTC).strftime('%Y-%m-%d %H:%M:%S')} UTC")
        
        fetcher = PttLinkFetcher()
        fetcher.connect_to_sheets()
        fetcher.update_sheet()
        
        end_time = time.time()
        duration = end_time - start_time
        logger.info(f"İşlem tamamlandı. Toplam süre: {duration:.2f} saniye")
        logger.info(f"Başarılı: {fetcher.success_count}, Hata: {fetcher.error_count}, Toplam: {fetcher.total_count}")
        
        print("\nPttAVM'den güncel linkler başarı ile çekilmiştir.")
        print(f"\nToplam: {fetcher.total_count} ürün işlendi, {fetcher.success_count} başarılı, {duration:.2f} saniye sürdü.")
        print("\nKuzey Uzun'dan sevgilerle :)")
        
    except Exception as e:
        logger.error(f"Program hatası: {str(e)}")
        logger.error(traceback.format_exc())
        raise
    finally:
        logger.info(f"Bitiş zamanı: {datetime.now(UTC).strftime('%Y-%m-%d %H:%M:%S')} UTC")

if __name__ == "__main__":
    main()