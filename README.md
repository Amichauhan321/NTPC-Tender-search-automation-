import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


# ------------------------
# CONFIG
# ------------------------

EXCEL_PATH = r"C:\Users\Dell\Downloads\Master Data Sheet.xlsx"
COLUMN_NAME = "Tender/ Ref No."


# ------------------------
# Create Chrome Driver
# ------------------------

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)  # browser manually close tak open rahega
    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(60)
    return driver


# ------------------------
# Manual captcha + search confirmation
# ------------------------

def wait_for_manual_confirmation(tender_id):
    input(f"👉 Tender {tender_id}: Captcha fill karke Search button manually click karo, "
          f"fir terminal me Enter dabao...")


# ------------------------
# Process one tender
# ------------------------

def process_tender(tender_id, driver):
    wait = WebDriverWait(driver, 40)

    try:
        # Step 1: Tender status page open
        driver.get("https://eprocurentpc.nic.in/nicgep/app?page=WebTenderStatusLists&service=page")

        # Step 2: Tender ID fill karo
        tender_box = wait.until(EC.presence_of_element_located((By.ID, "tenderId")))
        tender_box.clear()
        tender_box.send_keys(str(tender_id).strip())
        print(f"✅ Tender ID {tender_id} filled in search form")

        # Step 3: Captcha + Search aap manually karo
        wait_for_manual_confirmation(tender_id)

        # Step 4: Table wait karo
        wait.until(EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'list_table')]")))
        print("✅ Search results table loaded")

        # Step 5: Last row ke status ke neeche box icon par click karo (anchor wrapping <img>)
        status_box = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//table[contains(@class,'list_table')]//tr[last()]//td[last()]//a/img/parent::a")
        ))
        driver.execute_script("arguments[0].click();", status_box)
        print("✅ Status box link clicked")

        # Step 6: Stage summary link click karo (full XPath instead of partial link text)
        try:
            stage_link = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//a[contains(text(),'Click link to view the all stage summary Details')]")
                )
            )
            driver.execute_script("arguments[0].click();", stage_link)
            print("✅ All stage summary details link clicked")
        except TimeoutException:
            print("❌ Stage summary details link not found")

        print(f"🎯 Tender {tender_id} completed successfully till stage summary details")

    except TimeoutException as e:
        print(f"❌ Timeout error for {tender_id}: {e}")
    except Exception as e:
        print(f"⚠️ Unexpected error for {tender_id}: {e}")


# ------------------------
# Main
# ------------------------

def main():
    try:
        df = pd.read_excel(EXCEL_PATH)

        # Filter: Sirf NTPC + U column = Yes
        filtered = df[
            df.iloc[:, 5].astype(str).str.upper().str.contains("NTPC") &
            (df.iloc[:, 20].astype(str).str.strip().str.lower() == "yes")
        ]
        tender_ids = filtered[COLUMN_NAME].dropna().tolist()

    except Exception as e:
        print(f"⚠️ Could not read Excel: {e}")
        return

    print(f"Total NTPC + Green marked tenders found: {len(tender_ids)}")

    driver = create_driver()

    for tender_id in tender_ids:
        print(f"\n🔎 Checking Tender: {tender_id}")
        process_tender(tender_id, driver)

    input("👉 Sab kaam ho gaya. Browser manually check kar lo. Close karne ke liye yaha Enter dabao...")
    driver.quit()
    print("✅ All tenders processed, browser closed.")


if __name__ == "__main__":
    main()

