import pandas as pd
import time
import random
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.service import Service
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


class DoctoraliaPhoneExtractor:
    def __init__(self, excel_file_path, use_proxy=False, proxy_address=None):
        """
        Initialize the extractor

        Args:
            excel_file_path (str): Path to the Excel file
            use_proxy (bool): Whether to use a proxy
            proxy_address (str): Proxy address in format "ip:port"
        """
        self.excel_file_path = excel_file_path
        self.use_proxy = use_proxy
        self.proxy_address = proxy_address
        self.driver = None

    def setup_driver(self):
        """Set up Chrome WebDriver with options"""
        chrome_options = Options()

        # Basic options for stability
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        # User agent to appear more like a regular browser
        chrome_options.add_argument(
            "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"
        )

        # Proxy setup if needed
        if self.use_proxy and self.proxy_address:
            chrome_options.add_argument(f"--proxy-server={self.proxy_address}")
            logger.info(f"Using proxy: {self.proxy_address}")

        # Uncomment the line below to run headless (without opening browser window)
        # chrome_options.add_argument("--headless")

        # Point directly to installed chromedriver
        service = Service("/usr/local/bin/chromedriver")

        try:
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.driver.execute_script(
                "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
            )
            logger.info("Chrome WebDriver initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize WebDriver: {e}")
            raise

    def clean_phone(self, text):
        """
        Clean extracted text to get only the phone number and format it

        Args:
            text (str): Raw text containing phone number

        Returns:
            str: Cleaned and formatted phone number or None
        """
        digits = re.sub(r'\D', '', text)
        if len(digits) == 10:
            return f"{digits[:2]} {digits[2:6]} {digits[6:]}"
        return None

    def extract_phones(self, profile_url, row_index):
        """
        Extract up to two phone numbers from a single profile URL
        """
        extracted_phones = []
        try:
            logger.info(f"Processing row {row_index}: {profile_url}")
            self.driver.get(profile_url)

            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            time.sleep(random.uniform(2, 4))

            try:
                phone_containers = self.driver.find_elements(
                    By.CSS_SELECTOR, '[data-id="gdpr-show-number-block"]')
                logger.info(f"Found {len(phone_containers)} phone containers")

                for container_index, phone_container in enumerate(phone_containers):
                    try:
                        if len(extracted_phones) >= 2:
                            break

                        full_number_span = phone_container.find_element(
                            By.CSS_SELECTOR, 'span[data-id="shrinked-number"]')
                        partial_number = full_number_span.text.strip()

                        if "..." in partial_number:
                            logger.info(f"Container {container_index + 1}: Found partial number: {partial_number}. Attempting to reveal full number.")

                            show_phone_button = phone_container.find_element(
                                By.CSS_SELECTOR, '[data-id="show-phone-number-modal"]')
                            modal_target = show_phone_button.get_attribute('data-target')
                            logger.info(f"Modal target for container {container_index + 1}: {modal_target}")

                            self.driver.execute_script("arguments[0].click();", show_phone_button)
                            time.sleep(3)

                            try:
                                modal_data_id = None
                                if modal_target:
                                    match = re.search(r"data-id='([^']+)", modal_target)
                                    if match:
                                        modal_data_id = match.group(1)
                                        logger.info(f"Looking for modal with data-id: {modal_data_id}")

                                if modal_data_id:
                                    modal = WebDriverWait(self.driver, 10).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, f'[data-id="{modal_data_id}"]'))
                                    )
                                else:
                                    modal = WebDriverWait(self.driver, 10).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, '.modal[data-id*="phone"].show, .modal[data-id*="phone"]:not(.fade)'))
                                    )

                                logger.info(f"Modal appeared for container {container_index + 1}")
                                time.sleep(1)

                                phone_extracted = False
                                tel_links = modal.find_elements(By.CSS_SELECTOR, 'a[href^="tel:"]')
                                for link in tel_links:
                                    raw_phone = link.get_attribute('href').replace('tel:', '').strip()
                                    cleaned = self.clean_phone(raw_phone)
                                    if cleaned and cleaned not in extracted_phones:
                                        extracted_phones.append(cleaned)
                                        logger.info(f"Extracted phone from tel link in container {container_index + 1}: {cleaned}")
                                        phone_extracted = True
                                        break

                                if not phone_extracted:
                                    bold_elements = modal.find_elements(By.CSS_SELECTOR, 'b, strong')
                                    for elem in bold_elements:
                                        raw_text = elem.text.strip()
                                        cleaned = self.clean_phone(raw_text)
                                        if cleaned and cleaned not in extracted_phones:
                                            extracted_phones.append(cleaned)
                                            logger.info(f"Extracted phone from bold text in container {container_index + 1}: {cleaned}")
                                            phone_extracted = True
                                            break

                                if not phone_extracted:
                                    modal_text = modal.text
                                    matches = re.findall(r'\d{2}\s?\d{4}\s?\d{4}', modal_text)
                                    for match in matches:
                                        cleaned = self.clean_phone(match)
                                        if cleaned and cleaned not in extracted_phones:
                                            extracted_phones.append(cleaned)
                                            logger.info(f"Extracted phone from modal text in container {container_index + 1}: {cleaned}")
                                            phone_extracted = True
                                            break

                                try:
                                    close_button = modal.find_element(By.CSS_SELECTOR, '[data-dismiss="modal"], .close, button[aria-label="Close"]')
                                    self.driver.execute_script("arguments[0].click();", close_button)
                                    time.sleep(2)
                                except:
                                    self.driver.execute_script("arguments[0].style.display = 'none';", modal)
                                    try:
                                        backdrops = self.driver.find_elements(By.CSS_SELECTOR, '.modal-backdrop')
                                        for backdrop in backdrops:
                                            self.driver.execute_script("arguments[0].remove();", backdrop)
                                    except:
                                        pass
                                    time.sleep(2)

                                logger.info(f"Modal closed for container {container_index + 1}")

                            except TimeoutException:
                                logger.warning(f"Modal did not appear for container {container_index + 1} in row {row_index}")

                        else:
                            cleaned = self.clean_phone(partial_number)
                            if cleaned and cleaned not in extracted_phones:
                                extracted_phones.append(cleaned)
                                logger.info(f"Full phone number already visible in container {container_index + 1}: {cleaned}")

                        time.sleep(random.uniform(2, 3))

                    except NoSuchElementException as e:
                        logger.warning(f"Elements not found in container {container_index + 1} for row {row_index}: {e}")
                        continue
                    except Exception as e:
                        logger.warning(f"Error processing container {container_index + 1} for row {row_index}: {e}")
                        continue

            except TimeoutException:
                logger.warning(f"No phone containers found for row {row_index}")
            except Exception as e:
                logger.warning(f"Error finding phone containers for row {row_index}: {e}")

            logger.info(f"Total phones extracted for row {row_index}: {extracted_phones}")
            return extracted_phones[:2]

        except WebDriverException as e:
            logger.error(f"WebDriver error for row {row_index}: {e}")
            return []
        except Exception as e:
            logger.error(f"Unexpected error for row {row_index}: {e}")
            return []

    def process_excel_file(self, start_row=2, max_rows=None):
        """
        Process the Excel file and extract phone numbers
        """
        try:
            logger.info(f"Reading Excel file: {self.excel_file_path}")
            df = pd.read_excel(self.excel_file_path)

            if df.empty:
                logger.error("Excel file is empty")
                return

            if 'Phone1' not in df.columns:
                df['Phone1'] = ""
            if 'Phone2' not in df.columns:
                df['Phone2'] = ""

            logger.info(f"Found {len(df)} rows in Excel file")
            self.setup_driver()

            processed_count = 0
            start_index = start_row - 1
            end_index = min(len(df), start_index + max_rows) if max_rows else len(df)

            for index in range(start_index, end_index):
                try:
                    profile_url = df.iloc[index, 0]

                    if pd.isna(profile_url) or not profile_url:
                        logger.warning(f"No URL found in row {index + 1}")
                        df.loc[index, 'Phone1'] = "No URL"
                        continue

                    if not profile_url.startswith('http'):
                        profile_url = 'https://' + profile_url.lstrip('/')

                    phones = self.extract_phones(profile_url, index + 1)

                    df.loc[index, 'Phone1'] = phones[0] if phones else "No phone found"
                    df.loc[index, 'Phone2'] = phones[1] if len(phones) > 1 else ""

                    processed_count += 1
                    logger.info(f"Processed {processed_count}/{end_index - start_index} profiles")

                    if processed_count % 10 == 0:
                        df.to_excel(self.excel_file_path, index=False)
                        logger.info(f"Progress saved after {processed_count} records")

                    time.sleep(random.uniform(3, 6))

                except Exception as e:
                    logger.error(f"Error processing row {index + 1}: {e}")
                    df.loc[index, 'Phone1'] = f"Error: {str(e)}"
                    continue

            df.to_excel(self.excel_file_path, index=False)
            logger.info(f"Processing complete. Updated {processed_count} records.")

        except Exception as e:
            logger.error(f"Error processing Excel file: {e}")
            raise
        finally:
            if self.driver:
                self.driver.quit()
                logger.info("WebDriver closed")


def main():
    """Main function to run the phone extractor"""
    EXCEL_FILE_PATH = "/home/ubuntu/doctoralia/doctoralia.xlsx"
    USE_PROXY = False
    PROXY_ADDRESS = "proxy_ip:proxy_port"
    START_ROW = 2
    MAX_ROWS = 3900

    extractor = DoctoraliaPhoneExtractor(
        excel_file_path=EXCEL_FILE_PATH,
        use_proxy=USE_PROXY,
        proxy_address=PROXY_ADDRESS if USE_PROXY else None
    )

    try:
        extractor.process_excel_file(start_row=START_ROW, max_rows=MAX_ROWS)
        print("Phone extraction completed successfully!")
    except Exception as e:
        print(f"An error occurred: {e}")
        logger.error(f"Main execution error: {e}")


if __name__ == "__main__":
    main()
