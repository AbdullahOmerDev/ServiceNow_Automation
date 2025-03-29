# Script: Automated Tag Editor for Edge Browser
# Credits: Abdullah Omer (https://github.com/AbdullahOmerDev)
# Purpose: Automates the process of editing tags in a web application

import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# Kill any running Edge browser processes before starting
os.system("taskkill /F /IM msedge.exe /T")

# Path to Edge user profile - generic path, update as needed
edge_profile_path = r'C:\Users\[username]\AppData\Local\Microsoft\Edge\User Data'

# Configure Edge browser options
edge_options = webdriver.EdgeOptions()
edge_options.add_argument(f'user-data-dir={edge_profile_path}')
edge_options.add_argument('profile-directory=Default')

# Initialize the Edge WebDriver with configured options
driver = webdriver.Edge(options=edge_options)
wait = WebDriverWait(driver, 10)  # Set a 10-second timeout for waiting operations

# Navigate to the login page
login_url = 'https://your-instance.service-now.com/login.do'
driver.get(login_url)

try:
    # Find and click on the user account button to proceed with login
    next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-test-id='a.user@domain.com']")))
    next_button.click()
except (NoSuchElementException, TimeoutException) as e:
    # Handle login errors gracefully
    print("Error during login process:", e)
    driver.quit()
    exit()

# Define the URL for the incidents list page
incidents_url = 'https://your-instance.service-now.com/nav_to.do?uri=incident.list'

def find_element_safe(wait, by, value, timeout=10):
    """
    Safely find an element with explicit wait to handle potential timing issues.
    
    Args:
        wait: WebDriverWait object
        by: Selenium locator strategy (e.g., By.ID)
        value: Selector value for the strategy
        timeout: Maximum time to wait for element to appear
        
    Returns:
        The found element or None if not found
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

# Main loop for the automation process
while True:
    # Navigate to the incidents list page
    driver.get(incidents_url)
    print("Incidents list page opened")

    try:
        # Find the shadow DOM host element (macroponent)
        shadow_host = find_element_safe(wait, By.XPATH, "//*[starts-with(name(), 'macroponent')]")
        if not shadow_host:
            print("❌ Shadow DOM Host not found.")
            break

        # Find the iframe within the shadow DOM
        iframe = find_shadow_element(shadow_host, 'iframe')
        if not iframe:
            print("❌ iframe not found. Reloading...")
            continue  # Reload page if iframe is missing
        
        # Switch to the iframe context to interact with its elements
        driver.switch_to.frame(iframe)
        print("✅ Switched to iframe!")

        # Find and click on the tag name link
        tag_name = find_element_safe(wait, By.CSS_SELECTOR, "a.linked.formlink")
        if not tag_name:
            print("❌ Required elements not found.")
            break
        
        tag_name.click()
        print("✅ Clicked on the tag name.")
        time.sleep(2)

        # Select the second option in the "Viewable by" dropdown
        Viewable_by = find_element_safe(wait, By.XPATH, "//select[@id='label.viewable_by']/option[2]")
        if not Viewable_by:
            print("❌ Required elements not found.")
            break
        
        Viewable_by.click()
        print("✅ Clicked on the viewable by.") 
        time.sleep(2)

        # Find the group list field and enter "DD"
        group_list = find_element_safe(wait, By.ID, "sys_display.label.group_list")
        if not group_list:
            print("❌ Required elements not found.")
            break
        
        group_list.send_keys("DD")
        print("✅ Clicked on the group list.")
        time.sleep(2)

        # Click on the user list field 
        random_click = find_element_safe(wait, By.ID, "sys_display.label.user_list")
        if not random_click:
            print("❌ Required elements not found.")
            break
        
        random_click.click()
        print("✅ Clicked on the user list field")
        time.sleep(2)

        # Find and click the update button to save changes
        update_button = find_element_safe(wait, By.ID, "sysverb_update")
        if not update_button:
            print("❌ Required elements not found.")
            break
        
        update_button.click()
        print("✅ Clicked on the update button.")
        time.sleep(2)

    except Exception as e:
        # Handle any unexpected errors
        print(f"An error occurred: {e}")
        break

# Close the browser when finished or on error
print("❌ Closing the browser due to error or completion.")
driver.quit()