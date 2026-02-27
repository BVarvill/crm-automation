# crm_automation.py
import os
import pandas as pd
import re
from datetime import datetime
from zoneinfo import ZoneInfo
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchWindowException
import time

# === CONFIGURATION ===
CRM_URL    = os.environ.get("CRM_URL", "https://your-crm-url-here")
USERNAME   = os.environ.get("CRM_USERNAME", "your_username_here")
PASSWORD   = os.environ.get("CRM_PASSWORD", "")
EXCEL_PATH = os.environ.get("EXCEL_PATH", "crm_entry_sheet.xlsx")
TIMEOUT    = 10  # seconds for explicit waits

# Timezone mappings for all US states
USA_STATE_TIMEZONES = {
    'AL':'US/Central','AK':'US/Alaska','AZ':'US/Mountain','AR':'US/Central','CA':'US/Pacific',
    'CO':'US/Mountain','CT':'US/Eastern','DE':'US/Eastern','FL':'US/Eastern','GA':'US/Eastern',
    'HI':'US/Hawaii','ID':'US/Mountain','IL':'US/Central','IN':'US/Eastern','IA':'US/Central',
    'KS':'US/Central','KY':'US/Eastern','LA':'US/Central','ME':'US/Eastern','MD':'US/Eastern',
    'MA':'US/Eastern','MI':'US/Eastern','MN':'US/Central','MS':'US/Central','MO':'US/Central',
    'MT':'US/Mountain','NE':'US/Central','NV':'US/Pacific','NH':'US/Eastern','NJ':'US/Eastern',
    'NM':'US/Mountain','NY':'US/Eastern','NC':'US/Eastern','ND':'US/Central','OH':'US/Eastern',
    'OK':'US/Central','OR':'US/Pacific','PA':'US/Eastern','RI':'US/Eastern','SC':'US/Eastern',
    'SD':'US/Central','TN':'US/Central','TX':'US/Central','UT':'US/Mountain','VT':'US/Eastern',
    'VA':'US/Eastern','WA':'US/Pacific','WV':'US/Eastern','WI':'US/Central','WY':'US/Mountain'
}
COUNTRY_TIMEZONES = {
    'UNITED KINGDOM':'Europe/London','UK':'Europe/London',
    'FRANCE':'Europe/Paris','GERMANY':'Europe/Berlin'
}

# Nickname mappings
NICKNAMES = {
    'michael': ['mike'],
    'daniel': ['dan'],
    'joseph': ['joe'],
    'girard': ['gerry'],
    'gregory': ['greg'],
    'timothy': ['tim'],
    'benjamin':['ben'],
    'matthew':['matt'],
    'andrew':['andy'],
    'christina':['tina'],
    'jennifer':['jen'],
    'jennifer':['jenn'],
    'christopher':['chris'],
    'kimberly':['kim'],
    'jeffrey':['jeff'],
    'thomas':['tom'],
    'james':['jim'],
    'steven':['steve'],
    'norman':['norm'],
    'patrick':['pat'],
}

# === HELPERS ===
def get_timezone(country: str, state: str) -> str:
    c, s = country.strip().upper(), state.strip().upper()
    if c in ('USA','US','UNITED STATES') and s in USA_STATE_TIMEZONES:
        return USA_STATE_TIMEZONES[s]
    return COUNTRY_TIMEZONES.get(c, 'Europe/London')

# normalize for matching
def normalize_org(name: str) -> str:
    s = name.lower().strip()
    return s.strip()

# acceptable first name variants
def name_variants(first: str) -> list[str]:
    base = first.lower().strip()
    variants = [base]
    variants += NICKNAMES.get(base, [])
    return variants

# === SETUP ===
options = Options()
options.headless = False

driver  = webdriver.Firefox(options=options)
wait    = WebDriverWait(driver, TIMEOUT)
actions = ActionChains(driver)

try:
    # login
    print("Logging into CRM…")
    driver.get(CRM_URL)
    wait.until(EC.presence_of_element_located((By.ID,'username'))).send_keys(USERNAME)
    driver.find_element(By.ID,'userpass').send_keys(PASSWORD)
    driver.find_element(By.ID,'loginlink').click()
    wait.until(EC.presence_of_element_located((By.ID,'lms_lastname_search')))
    SEARCH_PAGE_URL = driver.current_url
    print("Login successful")

    df = pd.read_excel(EXCEL_PATH)
    print(f"Loaded {len(df)} rows from {EXCEL_PATH}")

    for idx, row in df.iterrows():
        first = str(row.get('firstName','')).strip()
        last = str(row.get('lastName','')).strip()
        parent_org = str(row.get('ParentOrg','')).strip()
        country = str(row.get('Country','')).strip()
        state = str(row.get('State','')).strip()
        paste_note = str(row.get('Paste_in_CRM','')).strip()

        if not (first and last and parent_org):
            print(f"Skipping row {idx}: incomplete data")
            continue
        print(f"\nRow {idx}: {first} {last} @ {parent_org}")

        # SEARCH LAST NAME
        print("Searching last name…")
        driver.get(SEARCH_PAGE_URL)
        ln = wait.until(EC.presence_of_element_located((By.ID,'lms_lastname_search')))
        ln.clear(); ln.send_keys(last + Keys.ENTER)

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'search_results')))
        try:
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'tr.sRes')))
            print("Results frame loaded and results present.")
        except TimeoutException:
            print(f"No results for {first} {last}")
            driver.switch_to.default_content()
            org = wait.until(EC.presence_of_element_located((By.ID,'lms_orgname_search')))
            print(f"Falling back to company search: '{normalize_org(parent_org)}'")
            org.clear(); org.send_keys(normalize_org(parent_org) + Keys.ENTER)
            time.sleep(2)
            continue

        # FILTER CANDIDATES
        print("Filtering candidates…")
        variants = name_variants(first)
        norm = normalize_org(parent_org)
        candidates = []
        for r in driver.find_elements(By.CSS_SELECTOR,'tr.sRes'):
            try:
                name_txt = r.find_element(By.CSS_SELECTOR,'td:nth-child(3) a.sLnk').text
                org_txt = r.find_element(By.CSS_SELECTOR,'td:nth-child(1) a.sLnk').text
            except NoSuchElementException:
                continue
            parts = name_txt.split(',',1)
            if len(parts)<2: continue
            nm = parts[1].split('-',1)[0].strip().lower()
            if not any(var in nm for var in variants): continue
            if norm in normalize_org(org_txt):
                candidates.append((r, name_txt, org_txt))

        if len(candidates)!=1:
            msg = "No match" if not candidates else "Multiple matches"
            print(f"{msg} for {first} {last} @ {parent_org}")
            driver.switch_to.default_content()
            org = wait.until(EC.presence_of_element_located((By.ID,'lms_orgname_search')))
            org.clear(); org.send_keys(norm + Keys.ENTER)
            time.sleep(2)
            continue

        # OPEN DETAIL (ensure double-click)
        row_elem, name_txt, org_txt = candidates[0]
        print(f"Match found: {name_txt} @ {org_txt}")
        link = row_elem.find_element(By.CSS_SELECTOR,'td:nth-child(3) a.sLnk')
        driver.execute_script("arguments[0].scrollIntoView();", link)
        print("Clicking entry to open detail…")
        link.click(); time.sleep(0.5)
        actions.move_to_element(link).double_click(link).perform()

        # SWITCH TO DETAIL FRAME
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME,'detailsframe2')))
        print("Detail frame loaded.")
        time.sleep(2)  # allow detail contents to render

        # NEW NOTE ENTRY - Save instead of cancel
        print("Locating New Note link…")
        note_link = None
        for a in driver.find_elements(By.CSS_SELECTOR,'a.bodylink'):
            href = a.get_attribute('href') or ''
            if 'newNote' in href and last.lower() in a.text.lower():
                note_link = a
                break
        if not note_link:
            print(f"New Note link not found for {first} {last}")
        else:
            print("Clicking New Note link…")
            note_link.click()
            try:
                editor = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,'form.wftRichTextEditor div.textBox')))
                print("Editor loaded.")
                driver.execute_script("arguments[0].innerText = arguments[1]", editor, paste_note)
                print("Note content injected.")
                save_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,'a.bodylink[href*=doRichTextSaveNote]')))
                print("Clicking Save Note…")
                save_link.click()
                print(f"Note saved for {first} {last}")
            except TimeoutException:
                print("Editor did not load in time.")

        # SCHEDULE CALL - Save instead of cancel
        print("Scheduling new call…")
        tz = get_timezone(country, state)
        now = datetime.now(ZoneInfo(tz)).replace(hour=9, minute=2, second=0, microsecond=0)
        uk = now.astimezone(ZoneInfo('Europe/London')).strftime('%H:%M')

        try:
            call = wait.until(EC.element_to_be_clickable((By.LINK_TEXT,'New Call')))
            call.click()
            main = driver.current_window_handle
            for h in driver.window_handles:
                if h!=main:
                    driver.switch_to.window(h)
                    break
            wait.until(EC.presence_of_element_located((By.NAME,'callReason')))
            time.sleep(1)
            driver.find_element(By.NAME,'callReason').send_keys('ICMA Intro Email')
            ct = driver.find_element(By.NAME,'callTime')
            ct.clear(); ct.send_keys(uk)
            driver.find_element(By.CSS_SELECTOR,"input[name='callPriority'][value='L']").click()
            time.sleep(1)
            save_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,'a.bodylink[href*=doSave]')))
            print("Clicking Save Call…")
            save_btn.click()
            print(f"Call saved for {first} {last}")
        except (TimeoutException, NoSuchWindowException) as e:
            print(f"Failed to schedule call for {first} {last}: {e}")
        finally:
            try:
                driver.switch_to.window(main)
            except:
                pass
            driver.switch_to.default_content()

    print("\nAll rows processed.")
finally:
    driver.quit()