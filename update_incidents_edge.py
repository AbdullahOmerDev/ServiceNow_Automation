"""
Automated Incident Updater
This script automates the process of updating incidents in a service management system.
It logs in, navigates to incident lists, and adds comments to incidents.

Credits: Abdullah Omer (https://github.com/AbdullahOmerDev)
"""

import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains

# Kill any existing Edge processes to prevent conflicts
os.system("taskkill /F /IM msedge.exe /T")

# Path to Edge profile - replace with a generic path that will be configured by the user
edge_profile_path = r'C:\Users\[USERNAME]\AppData\Local\Microsoft\Edge\User Data'  # User should update this

# Set up Edge options
edge_options = webdriver.EdgeOptions()
edge_options.add_argument(f'user-data-dir={edge_profile_path}')
edge_options.add_argument('profile-directory=Default')

# Initialize WebDriver with extended timeout for slower connections
driver = webdriver.Edge(options=edge_options)
wait = WebDriverWait(driver, 10)

# Navigate to the service portal login page
login_url = 'https://your-instance.service-now.com/login.do'  # Replace with actual login URL
driver.get(login_url)

try:
    # Look for the login element - using a generic selector reference
    # User should replace with their own username element selector
    next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-test-id='YOUR_USERNAME_HERE']")))
    next_button.click()
except (NoSuchElementException, TimeoutException) as e:
    print("Error during login process:", e)
    driver.quit()
    exit()

# URL for incidents list - filter parameters can be adjusted as needed
incidents_url = 'https://your-instance.service-now.com/nav_to.do?uri=incident_list.do%3Fsysparm_query%3Dactive%3Dtrue%5EORDERBYDESCsys_created_on%26sysparm_view%3Ddefault'  # Replace with actual incidents URL

def find_element_safe(wait, by, value, timeout=10):
    """Safe method to find an element with explicit wait.
    
    Args:
        wait: WebDriverWait object
        by: Type of selector (By.ID, By.CSS_SELECTOR, etc.)
        value: Selector value
        timeout: Maximum time to wait in seconds
        
    Returns:
        WebElement if found, None otherwise
    """
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except TimeoutException:
        return None

def find_shadow_element(host, selector):
    """Find element inside a shadow DOM.
    
    Args:
        host: Shadow DOM host element
        selector: CSS selector to find within the shadow DOM
        
    Returns:
        WebElement if found, None otherwise
    """
    try:
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", host)
        return shadow_root.find_element(By.CSS_SELECTOR, selector)
    except Exception:
        return None

# Main loop to continuously check and update incidents
while True:
    # Navigate to the incidents list page
    driver.get(incidents_url)
    print("Incidents list page opened")

    try:
        # Find the shadow DOM host element (for modern ServiceNow interfaces)
        shadow_host = find_element_safe(wait, By.XPATH, "//*[starts-with(name(), 'macroponent')]")
        if not shadow_host:
            print("❌ Shadow DOM Host not found.")
            break

        # Find the iframe within the shadow DOM
        iframe = find_shadow_element(shadow_host, 'iframe')
        if not iframe:
            print("❌ iframe not found. Reloading...")
            continue  # Reload page if iframe is missing
        
        # Switch to the iframe containing the incident list
        driver.switch_to.frame(iframe)
        print("✅ Switched to iframe!")

        # Find and click on the first incident link
        incident_url = find_element_safe(wait, By.CSS_SELECTOR, "a.linked.formlink")
        
        if not incident_url:
            print("❌ Required elements not found.")
            break
        
        incident_url.click()
        print("✅ Clicked on the incident link.")
        time.sleep(2)  # Allow page to load

        # Find the work notes text area
        work_notes = find_element_safe(wait, By.ID, "activity-stream-work_notes-textarea")
        
        if not work_notes:
            print("❌ Work notes textarea not found.")
            break
        
        # Add the update message to work notes - can be customized
        work_notes.send_keys("هل من تحديث؟")  # "Any updates?" in Arabic
        print("✅ Work notes updated.")
        time.sleep(2)  # Give time for input to register

        # Find and click the post button to submit the work notes
        post_button = find_element_safe(wait, By.CSS_SELECTOR, "button.btn.btn-default.activity-submit")
        
        if not post_button:
            print("❌ Post button not found.")
            break
        
        post_button.click()
        print("✅ Clicked on the post button.")
        time.sleep(2)  # Wait for submission to complete

    except Exception as e:
        print(f"An error occurred: {e}")
        break

# Close the browser when the script completes or encounters an error
print("❌ Closing the browser due to error or completion.")
driver.quit()