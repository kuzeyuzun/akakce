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
# Sürüm: 1.7
# Son Güncelleme: 2025-02-18 14:28:28 UTC
# Geliştirici: Kuzey Uzun
# Geliştirici Desteği: GitHub Copilot & Claude AI
# Mevcut Kullanıcı: seydauzun
############################################################

import os
import time
import logging
from datetime import datetime, UTC
from pathlib import Path
from typing import Optional
from dataclasses import dataclass

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

SPREADSHEET_ID = "1tMv7ElvqIBqmO6mqPCk_w-HprA9O02J4gVBGdXQ-4ns"
WORKSHEET_NAME = "Asorti_TV"
CREDENTIALS_FILE = "credentials.json"
WAIT_TIME = 10
DEBUG_MODE = True

@dataclass
class ProductData:
    """Ürün verisi için veri sınıfı"""
    ptt_link: str = ""

class PttLinkFetcher:
    def __init__(self):
        self.debug_mode = DEBUG_MODE
        self.setup_logging()
        
        # Google Sheets API yapılandırması
        self.scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        self.credentials_path = CREDENTIALS_FILE
        self.sheet_id = SPREADSHEET_ID
        self.worksheet_name = WORKSHEET_NAME
        
        # Chrome seçenekleri
        self.chrome_options = Options()
        self.chrome_options.add_argument('--headless')
        self.chrome_options.add_argument('--no-sandbox')
        self.chrome_options.add_argument('--disable-dev-shm-usage')
        self.chrome_options.add_argument('--disable-gpu')
        self.chrome_options.add_argument("--window-size=1920,1080")
        self.chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        self.chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        self.chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # İstatistikler
        self.success_count = 0
        self.error_count = 0
        self.total_count = 0
        
        # WebDriver
        self.driver = None

    def setup_logging(self):
        """Logging yapılandırmasını ayarlar"""
        try:
            log_dir = Path("logs")
            log_dir.mkdir(exist_ok=True)
            
            current_date = datetime.now().strftime("%Y%m%d")
            log_file = log_dir / f"ptt_finder_{current_date}.log"
            
            logging.basicConfig(
                level=logging.DEBUG if self.debug_mode else logging.INFO,
                format='[%(asctime)s] %(levelname)s: %(message)s',
                datefmt='%H:%M:%S',
                handlers=[
                    logging.FileHandler(log_file, encoding='utf-8'),
                    logging.StreamHandler()
                ]
            )
            
            self.logger = logging.getLogger(__name__)
            self.logger.info("Logging sistemi başlatıldı")
            
        except Exception as e:
            print(f"Logging kurulum hatası: {str(e)}")
            raise

    def print_status(self, message: str, level: str = 'info'):
        """Durum mesajlarını loglar"""
        if not hasattr(self, 'logger'):
            print(message)
            return

        if level == 'debug':
            self.logger.debug(message)
        elif level == 'warning':
            self.logger.warning(message)
        elif level == 'error':
            self.logger.error(message)
        else:
            self.logger.info(message)

    def connect_to_sheets(self):
        """Google Sheets bağlantısını kurar"""
        try:
            credentials = ServiceAccountCredentials.from_json_keyfile_name(
                self.credentials_path, self.scope)
            gc = gspread.authorize(credentials)
            self.sheet = gc.open_by_key(self.sheet_id).worksheet(self.worksheet_name)
            self.print_status("Google Sheets bağlantısı kuruldu")
            
            # İşlem başladı mesajı
            self.sheet.update(values=[["Lütfen Bekleyiniz..."]], range_name='O1')
            
        except Exception as e:
            self.print_status(f"Sheets bağlantı hatası: {str(e)}", 'error')
            raise

    def initialize_driver(self):
        """WebDriver'ı başlatır"""
        if not self.driver:
            self.driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=self.chrome_options
            )
            self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                'source': '''
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => undefined
                    })
                '''
            })

    def fetch_ptt_link(self, url: str) -> Optional[ProductData]:
        """Sayfadaki PTT AVM linkini bulur"""
        try:
            if not self.driver:
                self.initialize_driver()
            
            self.driver.get(url)
            wait = WebDriverWait(self.driver, WAIT_TIME)
            
            # Tüm satıcı linklerini bul
            try:
                # Satıcı tablosunu bekle
                wait.until(EC.presence_of_element_located((By.CLASS_NAME, "sellers_table")))
                
                # Tüm linkleri kontrol et
                links = self.driver.find_elements(By.TAG_NAME, "a")
                for link in links:
                    try:
                        href = link.get_attribute('href')
                        if href and 'pttavm.com' in href:
                            return ProductData(ptt_link=href)
                    except:
                        continue
                
                # Alternative: Satıcı hücrelerini kontrol et
                seller_cells = self.driver.find_elements(By.CSS_SELECTOR, "td.sl_v2")
                for cell in seller_cells:
                    try:
                        link = cell.find_element(By.TAG_NAME, "a")
                        href = link.get_attribute('href')
                        if href and 'pttavm.com' in href:
                            return ProductData(ptt_link=href)
                    except:
                        continue
                
            except Exception as e:
                self.print_status(f"Link arama hatası: {str(e)}", 'debug')
            
            return None
            
        except Exception as e:
            self.print_status(f"Sayfa işleme hatası ({url}): {str(e)}", 'error')
            return None

    def update_sheet(self):
        """PTT linklerini günceller"""
        try:
            # Tüm verileri al
            all_values = self.sheet.get_all_records()
            last_row = len(all_values) + 2
            
            self.print_status(f"Toplam {len(all_values)} satır işlenecek")
            
            # Her satır için işlem yap
            for row in range(2, last_row):
                try:
                    url = self.sheet.cell(row, 3).value  # C sütunu
                    if not url:
                        continue
                    
                    url = url.strip()
                    if not url.lower().startswith('https://'):
                        self.print_status(f"Satır {row}: Geçersiz URL", 'warning')
                        continue
                    
                    self.print_status(f"İşleniyor: Satır {row} - {url}")
                    
                    data = self.fetch_ptt_link(url)
                    if data and data.ptt_link:
                        # Sadece L sütununu güncelle (PTT Link)
                        self.sheet.update(values=[[data.ptt_link]], range_name=f'L{row}')
                        self.success_count += 1
                    else:
                        self.sheet.update(values=[[""]], range_name=f'L{row}')
                        self.error_count += 1
                    
                    time.sleep(1)
                    
                except Exception as e:
                    self.print_status(f"Satır {row} işleme hatası: {str(e)}", 'error')
                    self.error_count += 1
                finally:
                    self.total_count += 1
                    
        except Exception as e:
            self.print_status(f"Genel güncelleme hatası: {str(e)}", 'error')
            raise
        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
            try:
                self.sheet.update(values=[[""]], range_name='O1')
            except:
                pass

def main():
    start_time = time.time()
    try:
        print_status = logging.getLogger(__name__).info
        print_status("Program başlatılıyor...")
        print_status(f"Başlangıç zamanı: {datetime.now(UTC).strftime('%Y-%m-%d %H:%M:%S')} UTC")
        
        fetcher = PttLinkFetcher()
        fetcher.connect_to_sheets()
        fetcher.update_sheet()
        
        end_time = time.time()
        duration = end_time - start_time
        print_status(f"İşlem tamamlandı. Toplam süre: {duration:.2f} saniye")
        print_status(f"Başarılı: {fetcher.success_count}, Hata: {fetcher.error_count}, Toplam: {fetcher.total_count}")
        
        print("\nPttAVM'den güncel linkler başarı ile çekilmiştir.")
        print("\nKuzey Uzun'dan sevgilerle :)")
        
    except Exception as e:
        print_status(f"Program hatası: {str(e)}", 'error')
        raise
    finally:
        print_status(f"Bitiş zamanı: {datetime.now(UTC).strftime('%Y-%m-%d %H:%M:%S')} UTC")

if __name__ == "__main__":
    main()