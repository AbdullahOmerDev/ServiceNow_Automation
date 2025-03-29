"""
Automated Task Assignment Script for ServiceNow
Author: Abdullah Omer (https://github.com/AbdullahOmerDev)
Description: This script automates the assignment of tickets in ServiceNow based on assignment groups.
License: MIT
"""

import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# Kill any existing Chrome processes before starting
os.system("taskkill /F /IM chrome.exe /T")

# Configuration for Chrome profile
# REMOVED: Actual file path to personal Chrome profile
chrome_profile_path = r'[CHROME_PROFILE_PATH]'  # Replace with your Chrome profile path
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument(f'user-data-dir={chrome_profile_path}')
chrome_options.add_argument('profile-directory=Profile 2')

# Initialize the WebDriver with longer timeout for slow pages
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)  # 10-second timeout for finding elements

'''
# Login section - commented out, likely because the profile already has authentication cookies
# REMOVED: Actual login URL and credentials
login_url = '[LOGIN_URL]'
driver.get(login_url)

try:
    # Handle email input (if needed)
    try:
        email_input = wait.until(EC.presence_of_element_located((By.ID, 'i0116')))
        email_input.send_keys('[EMAIL]')  # Replace with actual email
        next_button = wait.until(EC.element_to_be_clickable((By.ID, 'idSIButton9')))
        next_button.click()
    except TimeoutException:
        print("Email step skipped, already on the password page.")

    # Handle password input
    password_input = wait.until(EC.presence_of_element_located((By.ID, 'i0118')))
    password_input.send_keys('[PASSWORD]')  # Replace with actual password
    
    # Submit login form
    sign_in_button = wait.until(EC.element_to_be_clickable((By.ID, 'idSIButton9')))
    sign_in_button.click()
except (NoSuchElementException, TimeoutException) as e:
    print("Error during login process:", e)
    driver.quit()
    exit()
'''

# URL for the incident list page that shows unassigned tickets
incidents_url = 'https://example.com/incident_list'  # Replace with actual URL

def find_element_safe(wait, by, value, timeout=10):
    """
    Safely find an element with explicit wait without raising exceptions.
    Returns the element if found, None otherwise.
    """
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except TimeoutException:
        return None

def find_shadow_element(host, selector):
    """
    Find an element inside a shadow DOM.
    
    Args:
        host: The shadow DOM host element
        selector: CSS selector to find within the shadow DOM
        
    Returns:
        The found element or None if not found
    """
    try:
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", host)
        return shadow_root.find_element(By.CSS_SELECTOR, selector)
    except Exception:
        return None

# Main automation loop
while True:
    # Navigate to the incidents page
    driver.get(incidents_url)
    print("Incidents list page opened")

    try:
        # Find the shadow DOM host element (modern web component)
        shadow_host = find_element_safe(wait, By.XPATH, "//*[starts-with(name(), 'macroponent')]")
        if not shadow_host:
            print("❌ Shadow DOM Host not found.")
            break

        # Find the iframe inside the shadow DOM
        iframe = find_shadow_element(shadow_host, 'iframe')
        if not iframe:
            print("❌ iframe not found. Reloading...")
            continue  # Reload page if iframe is missing
        
        # Switch to the iframe containing the incident list
        driver.switch_to.frame(iframe)
        print("✅ Switched to iframe!")

        # Find the implementer cell and assignment group cell in the first row
        implementer = find_element_safe(wait, By.CSS_SELECTOR, "tr[id^='row_incident_'] > *:nth-child(11)")
        assign_group_element = find_element_safe(wait, By.CSS_SELECTOR, "tr[id^='row_incident_'] > *:nth-child(10)")
        
        if not implementer or not assign_group_element:
            print("❌ Required elements (implementer, assign_group_element) not found.")
            break
        
        # Double-click on the implementer field to edit it
        ActionChains(driver).double_click(implementer).perform()
        print("✅ Double-clicked on implementer field")
        
        # Get the assignment group text to determine who should be assigned
        assign_group = assign_group_element.text
        
        # Mapping of assignment groups to specific implementers
        implementer_mapping = {
            'Group A': 'Implementer A',
            'Group B': 'Implementer B',
            'Group C': 'Implementer C',
            # Add more mappings as needed
        }
        
        # Determine which implementer to assign based on group, default to User
        implementer_text = implementer_mapping.get(assign_group, 'User')
        
        # Find the input field for assigning an implementer
        implementer_add = find_element_safe(wait, By.CSS_SELECTOR, "input#sys_display\\.LIST_EDIT_incident\\.assigned_to", timeout=5)
        
        if not implementer_add:
            print("❌ Implementer input field not found.")
            break
        
        # Enter the implementer name
        implementer_add.send_keys(implementer_text)
        time.sleep(2)  # Allow time for dropdown suggestions to appear
        
        # Find and click the confirm button
        implementer_add_button = find_element_safe(wait, By.CSS_SELECTOR, "a#cell_edit_ok", timeout=5)
        if not implementer_add_button:
            print("❌ Implementer add button not found.")
            break
        
        # Confirm the assignment
        implementer_add_button.click()
        print(f"✅ Implementer '{implementer_text}' added successfully!")
        
        # Wait before processing the next incident
        time.sleep(3)

    except Exception as e:
        print(f"An error occurred: {e}")
        break

# Clean up
print("❌ Closing the browser due to error or completion.")
driver.quit()