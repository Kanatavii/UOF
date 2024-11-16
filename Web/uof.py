import os
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def wait_for_uof_file_completion(download_path, filename='UOF出入库汇总表.xlsx', timeout=300):
    end_time = time.time() + timeout
    while time.time() < end_time:
        files = os.listdir(download_path)
        # Check if the file is fully downloaded and does not have a temporary extension
        if filename in files and not filename.endswith('.crdownload'):
            return os.path.join(download_path, filename)
        time.sleep(1)  # Check every second
    return None

def download_uof_file():
    download_path = os.path.expanduser('~\Server\JBC')

    # Create download path if it doesn't exist
    if not os.path.exists(download_path):
        os.makedirs(download_path)

    chrome_options = Options()
    prefs = {
        'download.default_directory': download_path,
        'download.prompt_for_download': False,
        'profile.default_content_settings.popups': 0,
        'directory_upgrade': True,
        'safebrowsing.enabled': True
    }
    chrome_options.add_experimental_option('prefs', prefs)

    # Remove headless for debugging purposes
    # chrome_options.add_argument('--headless')  # Run in headless mode
    chrome_options.add_argument('--disable-gpu')  # Disable GPU acceleration
    chrome_options.add_argument('--no-sandbox')  # Needed for Linux only
    chrome_options.add_argument('--window-size=1920x1080')  # Set window size

    driver = webdriver.Chrome(options=chrome_options)
    latest_uof_file = None  # Initialize variable

    try:
        # Initialize Chrome browser
        driver.get("https://quickconnect.to/")
        print("Opened QuickConnect website.")

        # Find the QuickConnect ID input box and enter information
        input_box = driver.find_element(By.ID, "input-id")
        input_box.send_keys("uof-jp")
        print("Entered QuickConnect ID.")

        # Use explicit wait to ensure the submit button is clickable
        wait = WebDriverWait(driver, 60)
        submit_button = wait.until(EC.element_to_be_clickable((By.ID, "input-submit")))
        submit_button.click()
        print("Submitted QuickConnect ID.")

        # Find the username input box and enter information
        username_box = wait.until(EC.presence_of_element_located((By.NAME, "username")))
        username_box.send_keys("anguri")
        login_button = driver.find_element(By.XPATH, "//div[contains(@class, 'login-btn-spinner-wrapper')]")
        login_button.click()
        print("Entered username and clicked login.")

        # Find the password input box and enter information
        password_box = wait.until(EC.presence_of_element_located((By.NAME, "current-password")))
        password_box.send_keys("uofjpA-1")
        login_button = driver.find_element(By.XPATH, "//div[contains(@class, 'login-btn-spinner-wrapper')]")
        login_button.click()
        print("Entered password and clicked login.")

        # Click on File Station
        first_element = wait.until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="sds-desktop-shortcut"]/div/li[1]/div[1]'))
        )
        driver.execute_script("arguments[0].click();", first_element)
        print("Clicked on File Station.")

        # Double-click on the UOF folder
        action = ActionChains(driver)
        uof_folder = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'UOF')]")))
        action.double_click(uof_folder).perform()
        print("Double-clicked on UOF folder.")

        # Double-click on the Transfer Data folder
        div_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(),'转运数据')]")))
        action.double_click(div_element).perform()
        print("Double-clicked on Transfer Data folder.")

        # Wait for file size elements to appear
        file_size_elements = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.x-grid3-cell-inner.x-grid3-col-filesize")))
        print(f"Found {len(file_size_elements)} file size elements.")

        # If no file size elements found, print message and exit
        if not file_size_elements:
            print("No file sizes found")
            return None

        # Desired filename to download
        desired_filename = "UOF出入库汇总表.xlsx"

        # Locate all elements containing the filename information
        file_elements = driver.find_elements(By.XPATH, f"//div[contains(text(), '{desired_filename}')]")

        # Filter out elements that contain "$" (temporary files)
        filtered_file_elements = [element for element in file_elements if "$" not in element.text]

        # Select the first matching file
        desired_file_element = filtered_file_elements[0] if filtered_file_elements else None

        # Download the specified file
        if desired_file_element:
            print(f"Attempting to download the desired file: {desired_filename}...")
            action.double_click(desired_file_element).perform()

            # Wait for download to complete
            latest_uof_file = wait_for_uof_file_completion(download_path)
            if latest_uof_file:
                print("Download Successful!")
            else:
                print("Download did not complete within the expected time.")
        else:
            print(f"File {desired_filename} not found.")

            # Pause execution for debugging purposes
            time.sleep(10)

    except TimeoutException as e:
        logging.exception("Timeout while waiting for an element.")
        print("Timeout occurred while waiting for an element: ", e)
    except Exception as e:
        logging.exception(f"Error downloading UOF file: {e}")
        print(f"Error downloading UOF file: {e}")
    finally:
        time.sleep(10)  # Wait before quitting to ensure download completes
        driver.quit()
        print("Driver quit.")

    return latest_uof_file  # Return the latest downloaded UOF file

if __name__ == "__main__":
    downloaded_file = download_uof_file()
    if downloaded_file:
        print(f"Downloaded file path: {downloaded_file}")
    else:
        print("UOF file download failed. Program exited.")
