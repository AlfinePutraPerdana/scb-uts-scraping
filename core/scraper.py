import time
import tempfile
import shutil
import atexit
from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge import service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service

class Scraper:
    def __init__(self, driver_path):
        self.driver = None
        self.driver_path = driver_path
        self.temp_user_data_dir = tempfile.mkdtemp()
        atexit.register(self.cleanup_temp_profile)

    def start_browser(self, url):
        # Set up Edge options with isolated profile
        options = EdgeOptions()
        options.add_argument(f"--user-data-dir={self.temp_user_data_dir}")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-extensions")
        
        # Use the Service class to specify the WebDriver path
        service = Service(self.driver_path)
        self.driver = webdriver.Edge(service=service, options=options)
        self.driver.get(url)
        print("Opened:", url)
        
    def browse_file(self):
        from PySide6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*.*)")
        if file_path:
            self.file_path_entry.setText(file_path)
            
    def wait_for_new_window(self):
        original_window = self.driver.current_window_handle
        for _ in range(20):
            time.sleep(1)
            if len(self.driver.window_handles) > 1:
                new_window = [w for w in self.driver.window_handles if w != original_window][0]
                self.driver.switch_to.window(new_window)
                return self.driver.title
        return None

    def scrape_data(self, fields):
        try:
            WebDriverWait(self.driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.NAME, "enqframe"))
            )
            print("Switched to 'enqframe'.")

            if self.driver.find_elements(By.ID, "info-table"):
                print("Info table found.")
                for label, selector in fields.items():
                    try:
                        if selector["type"] == "css":
                            element = self.driver.find_element(By.CSS_SELECTOR, selector["selector"])
                            print(f"{label}: {element.text.strip()}")
                        elif selector["type"] == "table_lookup":
                            table_id = selector["table_id"]
                            label_text = selector["label"]
                            rows = self.driver.find_elements(By.CSS_SELECTOR, f"#{table_id} tbody tr")
                            for row in rows:
                                cells = row.find_elements(By.TAG_NAME, "td")
                                if cells and cells[0].text.strip() == label_text:
                                    print(f"{label}: {cells[1].text.strip()}")
                                    break
                            else:
                                print(f"{label}: [NOT FOUND]")
                    except Exception as e:
                        print(f"{label}: [ERROR] {e}")
            else:
                print("No data table found.")
        except Exception as e:
            print("Error in scraping logic:", e)

    def close_browser(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            print("Browser closed.")
            
    def cleanup_temp_profile(self):
        try:
            shutil.rmtree(self.temp_user_data_dir, ignore_errors=True)
            print("Temporary user data directory cleaned up.")
        except Exception as e:
            print(f"Warning: Could not clean up temp profile: {e}")