import time
import os
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
import warnings
from selenium.webdriver.common.service import Service

# Credits: Abdullah Omer (https://github.com/AbdullahOmerDev)
# Script Purpose: Automates task assignment in ServiceNow based on assignment group mapping

# Suppress specific Selenium warnings
warnings.filterwarnings("ignore", category=UserWarning, module='selenium.webdriver.common.service')

# Kill any existing Edge processes
os.system("taskkill /F /IM msedge.exe /T")

# Set up logging
# Use a generic path for log file
log_directory = os.getenv('USERPROFILE', 'C:\\')
log_file_path = os.path.join(log_directory, "ServiceNow_Logs.txt")
logging.basicConfig(
    filename=log_file_path,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
# Disable logging from Selenium to reduce noise
logging.getLogger('selenium').setLevel(logging.CRITICAL)
# Set appropriate logging level for this script
logging.getLogger().setLevel(logging.INFO)

# Path to Edge profile - using a generic path
# Note: Update this path according to your environment
edge_profile_path = os.path.join(os.getenv('USERPROFILE', 'C:\\'), 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data')

# Set up Edge options
edge_options = webdriver.EdgeOptions()
edge_options.add_argument(f'user-data-dir={edge_profile_path}')
edge_options.add_argument('profile-directory=Default')
edge_options.add_argument("--headless=new")  # Run in headless mode
edge_options.add_argument("--window-size=1920,1080")  # Set a proper screen size

# Initialize WebDriver
driver = webdriver.Edge(options=edge_options)
wait = WebDriverWait(driver, 5)  # 5 second timeout for element interactions

# ServiceNow URL
# Note: Replace with your organization's ServiceNow URL
login_url = 'https://your_instance.service-now.com/'
driver.get(login_url)

try:
    # Look for the login element - replace with a generic selector in your environment
    # This example uses a data-test-id, but you'll need to update based on your login page structure
    next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[data-test-id='login.button']")))
    next_button.click()
except (NoSuchElementException, TimeoutException) as e:
    logging.error(f"Error during login process: {e}")
    driver.quit()
    exit()

# Open the incidents list page
# Note: Replace with your organization's specific ServiceNow incidents URL
incidents_url = 'https://your_instance.service-now.com/nav_to.do?uri=incident_list.do'

def find_element_safe(wait, by, value, timeout=10):
    """
    Safe method to find an element with explicit wait.
    
    Args:
        wait: WebDriverWait instance
        by: Type of selector (By.ID, By.CSS_SELECTOR, etc.)
        value: Selector value
        timeout: Maximum wait time in seconds
        
    Returns:
        WebElement if found, None otherwise
    """
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except TimeoutException:
        return None

def find_shadow_element(host, selector):
    """
    Find element inside a shadow DOM.
    
    Args:
        host: Shadow DOM host element
        selector: CSS selector for element within shadow DOM
        
    Returns:
        WebElement if found, None otherwise
    """
    try:
        shadow_root = driver.execute_script("return arguments[0].shadowRoot", host)
        return shadow_root.find_element(By.CSS_SELECTOR, selector)
    except Exception as e:
        return None

while True:
    # Navigate to the incidents page
    driver.get(incidents_url)

    try:
        # Find shadow DOM host element (ServiceNow uses shadow DOM extensively)
        shadow_host = find_element_safe(wait, By.XPATH, "//*[starts-with(name(), 'macroponent')]")
        if not shadow_host:
            logging.warning("Shadow host element not found, breaking loop")
            break

        # Find iframe within shadow DOM
        iframe = find_shadow_element(shadow_host, 'iframe')
        if not iframe:
            logging.warning("Iframe not found, retrying page load")
            continue  # Reload page if iframe is missing
        
        # Switch to the iframe containing the incident list
        driver.switch_to.frame(iframe)

        # Find the implementer and assignment group elements
        implementer = find_element_safe(wait, By.CSS_SELECTOR, "tr[id^='row_incident_'] > *:nth-child(11)")
        assign_group_element = find_element_safe(wait, By.CSS_SELECTOR, "tr[id^='row_incident_'] > *:nth-child(10)")
        if not implementer or not assign_group_element:
            logging.warning("Implementer or assignment group element not found")
            break
        
        # Double-click on the implementer field to activate edit mode
        ActionChains(driver).double_click(implementer).perform()
        
        # Get the assignment group text and map to the appropriate implementer
        assign_group = assign_group_element.text
        
        # Mapping of assignment groups to implementers
        # This dictionary maps each group to a specific person who should handle their tickets
        implementer_mapping = {
            'Group A': 'Implementer A',
            'Group B': 'Implementer B',
            'Group C': 'Implementer C',
            # Add more mappings as needed
        }
        
        # Get the appropriate implementer based on the assignment group
        # Default to a specific user if no mapping exists
        implementer_text = implementer_mapping.get(assign_group, 'Default User')
        
        # Find and fill the implementer input field
        implementer_add = find_element_safe(wait, By.CSS_SELECTOR, "input#sys_display\\.LIST_EDIT_incident\\.assigned_to", timeout=3)
        if not implementer_add:
            logging.warning("Implementer input field not found")
            break
        
        # Enter the implementer name
        implementer_add.send_keys(implementer_text)
        time.sleep(2)  # Allow time for dropdown suggestions to appear
        
        # Click the OK button to confirm the assignment
        implementer_add_button = find_element_safe(wait, By.CSS_SELECTOR, "a#cell_edit_ok", timeout=5)
        if not implementer_add_button:
            logging.warning("OK button not found for implementer assignment")
            break
        
        # Complete the assignment
        implementer_add_button.click()
        logging.info(f"Implementer '{implementer_text}' added successfully to '{assign_group}'.")
        
        # Wait before processing the next entry
        time.sleep(3)

    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        break

# Close the browser at the end
logging.info("Closing the browser due to error or completion.")
driver.quit()