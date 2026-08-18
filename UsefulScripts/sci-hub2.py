#!/usr/bin/env python3
"""
Sci-Hub Article Downloader - Working Version
Properly handles cookies, URL cleaning, and content verification
"""

import pandas as pd
import requests
import time
import os
import re
import urllib3
from urllib.parse import quote, urljoin, urlparse, urlunparse
from pathlib import Path
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, WebDriverException

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class SciHubDownloader:
    def __init__(self, base_url="https://sci-hub.su", download_dir="downloads", delay=5):
        self.base_urls = [
            "https://sci-hub.su",
            "https://sci-hub.se", 
            "https://sci-hub.st",
            "https://sci-hub.ru",
            "https://sci-hub.wf",
            "https://sci-hub.ren",
            "https://sci-hub.mksa.top",
            "https://sci-hub.hkvisa.net",
            "https://sci-hub.shop",
        ]
        self.base_url = base_url.rstrip('/')
        self.download_dir = Path(download_dir)
        self.delay = delay
        self.driver = None
        self.session = None
        
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self.downloaded = 0
        self.failed = 0
        self.skipped = 0

    def init_selenium(self):
        """Initialize headless Chrome with proper settings."""
        if self.driver:
            return True
            
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # User agent
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.0')
        
        # Disable images
        prefs = {
            'profile.managed_default_content_settings.images': 2,
            'download.prompt_for_download': False,
            'download.default_directory': str(self.download_dir.absolute()),
        }
        chrome_options.add_experimental_option('prefs', prefs)
        
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            
            # Create session with same headers
            self.session = requests.Session()
            self.session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Referer': 'https://sci-hub.su/',
                'Connection': 'keep-alive',
            })
            
            print("  ✓ Browser initialized")
            return True
            
        except Exception as e:
            print(f"  ✗ Browser failed: {e}")
            return False

    def close_selenium(self):
        if self.driver:
            self.driver.quit()
            self.driver = None

    def sanitize_filename(self, filename):
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename[:200]

    def clean_pdf_url(self, url):
        """Remove fragments and tracking from PDF URL."""
        if not url:
            return None
        
        # Remove URL fragments (#view=FitH, etc)
        url = url.split('#')[0]
        
        # Remove tracking parameters
        parsed = urlparse(url)
        
        # Reject ad/tracking URLs
        bad_domains = ['google', 'criteo', 'doubleclick', 'facebook', 'adtraffic', 'syncframe']
        if any(bad in parsed.netloc.lower() for bad in bad_domains):
            return None
        
        # Reconstruct without query params if from sci-hub
        if 'sci-hub' in parsed.netloc or 'bban' in parsed.netloc:
            # Keep the URL clean
            clean = urlunparse((
                parsed.scheme or 'https',
                parsed.netloc,
                parsed.path,
                '', '', ''
            ))
            return clean
        
        return url

    def get_pdf_url_selenium(self, sci_hub_url):
        """Get PDF URL using Selenium with proper waiting."""
        try:
            self.driver.get(sci_hub_url)
            
            # Wait for page to load
            time.sleep(3)
            
            # Try multiple selectors
            selectors = [
                ('iframe#pdf', 'src'),
                ('iframe', 'src'),
                ('embed', 'src'),
                ('#pdf', 'src'),
                ('.pdf', 'src'),
                ('button[onclick*="location"]', 'onclick'),
            ]
            
            for selector, attr in selectors:
                try:
                    elem = self.driver.find_element(By.CSS_SELECTOR, selector)
                    if attr == 'onclick':
                        onclick = elem.get_attribute('onclick')
                        match = re.search(r'location\.href=["\']([^"\']+)["\']', onclick)
                        if match:
                            return match.group(1)
                    else:
                        src = elem.get_attribute(attr)
                        if src and len(src) > 10:
                            return src
                except:
                    continue
            
            # Parse page source
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            
            # Look for any element with PDF link
            for tag in ['iframe', 'embed', 'a', 'object']:
                for elem in soup.find_all(tag):
                    src = elem.get('src') or elem.get('data') or elem.get('href')
                    if src and ('.pdf' in src or '/pdf/' in src or len(src) > 30):
                        return src
            
            # JavaScript variables
            try:
                js_urls = self.driver.execute_script("""
                    var urls = [];
                    for (var key in window) {
                        if (window[key] && typeof window[key] === 'string') {
                            var val = window[key];
                            if ((val.includes('.pdf') || val.includes('/pdf/')) && val.startsWith('http')) {
                                urls.push(val);
                            }
                        }
                    }
                    return urls;
                """)
                if js_urls:
                    return js_urls[0]
            except:
                pass
                
        except Exception as e:
            print(f"  Selenium error: {e}")
        
        return None

    def download_with_selenium(self, pdf_url, filepath):
        """Download PDF using Selenium's network or direct request with cookies."""
        try:
            # Clean URL
            pdf_url = self.clean_pdf_url(pdf_url)
            if not pdf_url:
                return False, "Invalid URL (ads/tracking)"
            
            print(f"  Clean URL: {pdf_url[:70]}...")
            
            # Get cookies from Selenium
            cookies = self.driver.get_cookies()
            cookie_dict = {c['name']: c['value'] for c in cookies}
            
            # Download with session + cookies
            headers = {
                'User-Agent': self.driver.execute_script("return navigator.userAgent;"),
                'Accept': 'application/pdf,application/x-pdf,application/octet-stream,*/*',
                'Referer': self.driver.current_url,
            }
            
            response = self.session.get(
                pdf_url, 
                headers=headers, 
                cookies=cookie_dict,
                timeout=60, 
                stream=True, 
                verify=False,
                allow_redirects=True
            )
            
            # Check if we got redirected to HTML
            content_type = response.headers.get('Content-Type', '').lower()
            
            # Save and verify
            content = response.content
            
            if len(content) < 1000:
                return False, f"File too small ({len(content)} bytes)"
            
            # Check if it's actually PDF
            if content[:4] == b'%PDF':
                with open(filepath, 'wb') as f:
                    f.write(content)
                return True, f"PDF ({len(content):,} bytes)"
            
            # Check for PDF signature later in file
            if b'%PDF' in content[:1000]:
                with open(filepath, 'wb') as f:
                    f.write(content)
                return True, f"PDF found at offset ({len(content):,} bytes)"
            
            # Got HTML instead
            if b'<html' in content[:100].lower():
                # Try to extract PDF from this HTML
                soup = BeautifulSoup(content, 'html.parser')
                for iframe in soup.find_all('iframe'):
                    src = iframe.get('src')
                    if src and '.pdf' in src:
                        return self.download_with_selenium(src, filepath)
                return False, "Got HTML page instead of PDF"
            
            # Unknown content
            return False, f"Unknown content type: {content[:20]}"
            
        except Exception as e:
            return False, str(e)[:50]

    def download_paper(self, doi, serial_no):
        """Main download method."""
        if pd.isna(doi) or not str(doi).strip():
            return False, "Empty DOI", None
        
        doi = str(doi).strip()
        safe_serial = self.sanitize_filename(str(serial_no))
        filename = f"{safe_serial}.pdf"
        filepath = self.download_dir / filename
        
        if filepath.exists():
            self.skipped += 1
            return True, "Already exists", str(filepath)
        
        # Try each mirror
        mirrors = [self.base_url] + [u for u in self.base_urls if u != self.base_url]
        
        for mirror in mirrors:
            encoded_doi = quote(doi, safe='')
            sci_hub_url = f"{mirror}/{encoded_doi}"
            
            print(f"  Mirror: {mirror.split('//')[1]}")
            
            try:
                # Get PDF URL via Selenium
                pdf_url = self.get_pdf_url_selenium(sci_hub_url)
                
                if not pdf_url:
                    continue
                
                print(f"  Found: {pdf_url[:60]}...")
                
                # Download
                success, message = self.download_with_selenium(pdf_url, filepath)
                
                if success:
                    self.downloaded += 1
                    return True, message, str(filepath)
                else:
                    print(f"  ✗ {message}")
                    
            except Exception as e:
                print(f"  ✗ Error: {e}")
                continue
        
        self.failed += 1
        return False, "All mirrors failed", None

    def process_excel(self, excel_path, doi_column="DOI", serial_column="S.No.", 
                      sheet_name=0, start_row=0, end_row=None):
        """Process Excel file."""
        print(f"\n{'='*70}")
        print("SCI-HUB DOWNLOADER - FIXED VERSION")
        print(f"{'='*70}\n")
        
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            print(f"Excel: {len(df)} rows | Columns: {list(df.columns)}\n")
        except Exception as e:
            print(f"Error reading Excel: {e}")
            return
        
        if doi_column not in df.columns or serial_column not in df.columns:
            print(f"Error: Columns not found. Available: {list(df.columns)}")
            return
        
        if end_row is None:
            end_row = len(df)
        df = df.iloc[start_row:end_row]
        
        total = len(df)
        print(f"Processing: {total} papers | Directory: {self.download_dir.absolute()}")
        print(f"Delay: {self.delay}s")
        print(f"\n{'='*70}\n")
        
        # Initialize browser
        if not self.init_selenium():
            print("Failed to initialize browser. Exiting.")
            return
        
        try:
            for idx, row in df.iterrows():
                actual_row = start_row + idx + 2
                doi = row[doi_column]
                serial_no = row[serial_column]
                
                print(f"[{idx+1}/{total}] Row {actual_row} | S.No: {serial_no}")
                print(f"  DOI: {doi}")
                
                success, message, _ = self.download_paper(doi, serial_no)
                status = "✓" if success else "✗"
                print(f"  {status} {message}\n")
                
                if idx < total - 1:
                    time.sleep(self.delay)
        finally:
            self.close_selenium()
        
        print(f"\n{'='*70}")
        print(f"SUMMARY: {self.downloaded} OK | {self.skipped} Skip | {self.failed} Fail")
        print(f"{'='*70}")


def main():
    import sys
    
    CONFIG = {
        'excel_file': sys.argv[1] if len(sys.argv) > 1 else 'scopus_data.xlsx',
        'doi_column': sys.argv[2] if len(sys.argv) > 2 else 'DOI',
        'serial_column': sys.argv[3] if len(sys.argv) > 3 else 'S.No.',
        'download_dir': 'downloaded_papers',
        'delay': 5,
    }
    
    downloader = SciHubDownloader(
        download_dir=CONFIG['download_dir'], 
        delay=CONFIG['delay']
    )
    
    downloader.process_excel(
        excel_path=CONFIG['excel_file'],
        doi_column=CONFIG['doi_column'],
        serial_column=CONFIG['serial_column']
    )


if __name__ == "__main__":
    main()
