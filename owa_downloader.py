#!/usr/bin/env python3
import os
import time
import argparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager

def init_driver(download_dir):
    options = EdgeOptions()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_argument("--safebrowsing-disable-download-protection")
    options.add_argument("--log-level=3")
    options.add_argument("--disable-logging")
    options.add_argument("--disable-dev-shm-usage")
    
    # Suppress DevTools listening message
    os.environ['WDM_LOG_LEVEL'] = '0'
    os.environ['WDM_PRINT_FIRST_LINE'] = 'False'
    
    service = EdgeService(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=options)
    driver.maximize_window()
    driver.execute_cdp_cmd("Page.setDownloadBehavior", {
        "behavior": "allow",
        "downloadPath": download_dir
    })
    return driver

def wait_for_nav(driver, timeout=60):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='navigation']"))
    )

def login_manual(driver, owa_url):
    driver.get(owa_url)
    print("\n1) Please log in to Outlook Web Access in the browser window.")
    input("   Once signed in and you see your mailbox, press ENTER here…")
    wait_for_nav(driver)
    print("   Main mailbox detected.")

def open_shared_manual(driver):
    original = driver.current_window_handle
    print("\n2) Please open the shared mailbox in the browser:")
    print("     • Click your profile icon")
    print("     • Choose “Open another mailbox”")
    print("     • Enter the shared mailbox address and confirm")
    input("   After the shared mailbox opens in its own tab, press ENTER here…")
    for h in driver.window_handles:
        if h != original:
            driver.switch_to.window(h)
            break
    wait_for_nav(driver)
    print("   Switched into shared mailbox.")

def prepare_manual_view(driver):
    print("\n3) In the browser:")
    print("     • Disable conversation view: View → Show as → Messages")
    print("     • Navigate into the folder you want")
    print("     • Sort the list Oldest → Newest")
    input("   When that’s done, press ENTER here to start downloading…")

def download_and_delete_emails(driver, output_dir):
    wait = WebDriverWait(driver, 20)
    overflow_css = "button[aria-label='Weitere Aktionen'],button[aria-label='More actions']"
    processed = 0

    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        overflows = driver.find_elements(By.CSS_SELECTOR, overflow_css)
        toggles = driver.find_elements(By.CSS_SELECTOR, "button i[data-icon-name='ChevronRightSmall']")
        if not overflows:
            if toggles:
                print(f"{len(toggles)} collapsed threads remain—skipping them.")
            break

        try:
            btn = overflows[0]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            btn.click()
            time.sleep(0.5)

            download_found = False
            try:
                sw = WebDriverWait(driver, 2)
                dl = sw.until(EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[@role='menuitem' and ( .//span[contains(text(),'Herunterladen')] or .//span[contains(text(),'Download')] )]"
                )))
                dl.click()
                time.sleep(0.3)
                eml = sw.until(EC.element_to_be_clickable((
                    By.XPATH,
                    "//button[@role='menuitem' and ( .//span[contains(text(),'Download as EML')] or .//span[contains(text(),'EML')] )]"
                )))
                eml.click()
                time.sleep(1)
                download_found = True
            except TimeoutException:
                download_found = False

            btn2 = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, overflow_css))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn2)
            btn2.click()

            delete_btn = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                "//button[@role='menuitem' and ( .//span[contains(text(),'Löschen')] or .//span[contains(text(),'Delete')] )]"
            )))
            delete_btn.click()

            processed += 1
            verb = "saved & deleted" if download_found else "skipped download & deleted"
            print(f"{processed} {verb}")
            time.sleep(1)

        except StaleElementReferenceException:
            print("Element went stale—retrying this item…")
            continue
        except Exception as e:
            print(f"Unexpected error on item {processed+1}: {e}")
            break

    print(f"Done — {processed} messages processed, collapsed threads skipped.")

def main():
    p = argparse.ArgumentParser(description="Download & delete OWA emails (attachments left in .eml)")
    p.add_argument("--owa-url", default="https://outlook.office365.com/owa", help="Your Outlook Web Access URL")
    p.add_argument("--output-dir", default="downloads", help="Directory to save .eml files")
    p.add_argument("--mailbox", choices=["own", "shared"], default="own", help="Choose which mailbox to use: own or shared")
    args = p.parse_args()

    driver = None
    try:
        os.makedirs(args.output_dir, exist_ok=True)
        abs_output_dir = os.path.abspath(args.output_dir)
        print(f"\nOWA Backup Tool starting...")
        print(f"Files will be saved to: {abs_output_dir}")

        driver = init_driver(abs_output_dir)

        login_manual(driver, args.owa_url)
        if args.mailbox == "shared":
            open_shared_manual(driver)
        prepare_manual_view(driver)
        download_and_delete_emails(driver, args.output_dir)

        print("\nBackup completed successfully!")
        return 0

    except KeyboardInterrupt:
        print("\n\nUser cancelled. Closing browser...")
        return 1

    except Exception as e:
        print(f"\nError: {e}")
        print("\nThe program encountered an error. Please check your internet connection and try again.")
        input("Press ENTER to exit...")
        return 1

    finally:
        if driver:
            try:
                driver.quit()
                print("Browser closed.")
            except Exception as e:
                print(f"Error closing browser: {e}")
if __name__ == "__main__":
    import sys
    sys.exit(main())
