import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# Author: Abdullah Omer
# GitHub: https://github.com/AbdullahOmerDev
# Description: This script automates the resolution of incidents in a ticketing system
# by changing their state to "Resolved" and adding resolution information.

# Kill any existing Edge processes to ensure clean start
os.system("taskkill /F /IM msedge.exe /T")

# Path to Edge profile - generic path that should be modified by user
edge_profile_path = r'C:\Users\[username]\AppData\Local\Microsoft\Edge\User Data'

# Set up Edge options
edge_options = webdriver.EdgeOptions()
edge_options.add_argument(f'user-data-dir={edge_profile_path}')
edge_options.add_argument('profile-directory=Default')

# Initialize WebDriver with 10-second wait timeout
driver = webdriver.Edge(options=edge_options)
wait = WebDriverWait(driver, 10)

# Navigate to the service portal login page
login_url = 'https://your-instance.service-now.com/login.do'
driver.get(login_url)

try:
    # Look for and click the login option
    # Note: Replace with your own login selector or method
    next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-test-id='user@domain.com']")))
    next_button.click()
except (NoSuchElementException, TimeoutException) as e:
    print("Error during login process:", e)
    driver.quit()
    exit()

# URL for incidents that are held and assigned to a specific user
# Remove personal identifiers from URL query parameters
incidents_url = 'https://your-instance.service-now.com/nav_to.do?uri=incident_list.do%3Fsysparm_query%3Dactive%3Dtrue%5Eassigned_to%3Djavascript:gs.getUserID()%5EORDERBYDESCsys_created_on%26sysparm_view%3Dessentials'

def find_element_safe(wait, by, value, timeout=10):
    """
    Safely find an element with explicit wait.
    Returns None if element is not found within the timeout period.
    """
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except TimeoutException:
        return None

def find_shadow_element(host, selector):
    """
    Find an element inside a shadow DOM.
    Returns None if element cannot be found.
    """
    try:
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", host)
        return shadow_root.find_element(By.CSS_SELECTOR, selector)
    except Exception:
        return None

def click_element_safe(wait, element):
    """
    Clicks an element using WebDriverWait with an explicit wait for clickability.
    """
    wait.until(EC.element_to_be_clickable(element)).click()

def select_tab_by_text(wait, tab_text):
    """
    Selects a tab by its text content.
    """
    tab_headers = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "span.tab_header")))
    for tab_header in tab_headers:
        try:
            tab_element = tab_header.find_element(By.CSS_SELECTOR, "span.tabs2_tab")
            
            if tab_text in tab_element.text:
                click_element_safe(wait, tab_element)
                return  # Tab found and clicked, exit the function
        except Exception as e:
            print(f"Error processing tab: {e}")
    raise Exception(f"Tab with text '{tab_text}' not found.")

# Main automation loop
while True:
    # Navigate to incidents list page
    driver.get(incidents_url)
    print("Incidents list page opened")

    try:
        # Find shadow DOM host element (ServiceNow UI framework uses shadow DOM)
        shadow_host = find_element_safe(wait, By.XPATH, "//*[starts-with(name(), 'macroponent')]")
        if not shadow_host:
            print("❌ Shadow DOM Host not found.")
            break

        # Access iframe within shadow DOM
        iframe = find_shadow_element(shadow_host, 'iframe')
        if not iframe:
            print("❌ iframe not found. Reloading...")
            continue  # Reload page if iframe is missing
        
        # Switch to iframe containing the incident list
        driver.switch_to.frame(iframe)
        print("✅ Switched to iframe!")

        # Find and click on the first incident in the list
        incident_url = find_element_safe(wait, By.CSS_SELECTOR, "a.linked.formlink")
        if not incident_url:
            print("❌ No incidents found to process.")
            break
        
        incident_url.click()
        print("✅ Clicked on incident link.")
        time.sleep(2)

        # Find incident state dropdown
        set_state = find_element_safe(wait, By.CSS_SELECTOR, "select#incident\\.state")
        if not set_state:
            print("❌ Incident state dropdown not found.")
            break

        # Change incident state to "Resolved"
        select_state = Select(set_state)
        select_state.select_by_visible_text("Resolved")
        print("✅ Changed state to Resolved.")
        time.sleep(2)

        # Switch to the Resolution Information tab
        select_tab_by_text(wait, "Resolution Information")
        print("✅ Switched to Resolution Information tab.")
        time.sleep(2)

        # Select resolution code
        set_code = find_element_safe(wait, By.ID, "incident.close_code")
        if not set_code:
            print("❌ Resolution code dropdown not found.")
            break

        select_code = Select(set_code)
        select_code.select_by_visible_text("Solution provided")
        print("✅ Selected 'Solution provided' resolution code.")
        time.sleep(2)

        # Add resolution notes
        Resolution_note = find_element_safe(wait, By.CSS_SELECTOR, "textarea#incident\\.close_notes")
        if not Resolution_note:
            print("❌ Resolution notes field not found.")
            break

        # Add standardized closing note
        Resolution_note.send_keys("لعدم رد العميل لاكثر من 3 مرات يرجي اغلاق التذكرة")
        print("✅ Added resolution notes.")
        time.sleep(2)

        # Click update button to save changes
        update_button = find_element_safe(wait, By.ID, "sysverb_update")
        if not update_button:
            print("❌ Update button not found.")
            break
        
        update_button.click()
        print("✅ Clicked update button, incident resolved successfully.")
        time.sleep(2)

    except Exception as e:
        print(f"An error occurred: {e}")
        break

# Close the browser when finished
print("❌ Script completed. Closing browser.")
driver.quit()