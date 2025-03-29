# ServiceNow Automation Tools

A collection of Selenium-based automation scripts for ServiceNow ticketing system tasks. These scripts automate routine ServiceNow operations such as assigning tickets, updating incidents, resolving tickets, and editing tags.

## Author
**Abdullah Omer** - [AbdullahOmerDev](https://github.com/AbdullahOmerDev)

## Overview
This project contains several automation scripts that help ServiceNow administrators and technicians save time on repetitive tasks. The scripts use Selenium WebDriver to interact with the ServiceNow web interface.

## Features
- **Automatic Ticket Assignment**: Assign tickets to implementers based on assignment group mappings
- **Incident Resolution**: Automatically resolve incidents with customizable resolution notes
- **Incident Updates**: Add work notes to incidents to prompt for updates
- **Tag Editing**: Automate the process of editing tags with predefined values

## Scripts

### Ticket Assignment
- `assign_task_chrome.py` - Assigns tickets using Chrome browser
- `assign_task_edge.py` - Assigns tickets using Edge browser
- `assign_task_edge.bat` - Batch file to run the Edge assignment script in minimized mode

### Incident Management
- `resolve_incidents_edge.py` - Automatically resolves incidents with standard closing notes
- `update_incidents_edge.py` - Adds follow-up notes to incidents asking for updates

### Tag Management
- `edit_tag_edge.py` - Automates the process of editing tags in ServiceNow

## Requirements
- Python 3.x
- Selenium WebDriver
- Appropriate WebDriver executables:
  - ChromeDriver for Chrome-based scripts
  - EdgeDriver for Edge-based scripts
- A valid ServiceNow account with appropriate permissions

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/AbdullahOmerDev/ServiceNow_Automate.git
   ```

2. Install the required dependencies:
   ```
   pip install selenium
   ```

3. Download appropriate WebDriver executables:
   - [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/downloads)
   - [EdgeDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/)

4. Place WebDriver executables in the appropriate directories:
   - `chromedriver_win64/`
   - `edgedriver_win64/`
   - `edgedriver_arm64/`

## Configuration

Before running the scripts, you need to update the following in each script:

1. Browser profile paths - Update to your specific profile path
2. ServiceNow instance URLs - Replace with your organization's ServiceNow URLs
3. Implementer mappings - Configure the mapping dictionary to match your organization's structure

Example configuration (from `assign_task_edge.py`):
```python
# Path to Edge profile - using a generic path
edge_profile_path = os.path.join(os.getenv('USERPROFILE', 'C:\\'), 'AppData', 'Local', 'Microsoft', 'Edge', 'User Data')

# ServiceNow URL
login_url = 'https://your_instance.service-now.com/'

# Mapping of assignment groups to implementers
implementer_mapping = {
    'Group A': 'Implementer A',
    'Group B': 'Implementer B',
    'Group C': 'Implementer C',
    # Add more mappings as needed
}
```

## Usage

### Running Chrome Scripts
```
python assign_task_chrome.py
```

### Running Edge Scripts
```
python assign_task_edge.py
```
Or use the batch file to run in minimized mode:
```
assign_task_edge.bat
```

## Automation Logic

The scripts follow this general workflow:
1. Launch browser with pre-configured profile (assumes you're already authenticated)
2. Navigate to the ServiceNow incidents/tasks page
3. Find and interact with elements through shadow DOM when necessary
4. Perform the specific automation task (assign, update, resolve, edit)
5. Continue to the next item or exit upon completion

## Notes

- These scripts assume you have an active ServiceNow session in your browser profile
- Error handling is implemented to manage common issues like missing elements
- Scripts use shadow DOM traversal for modern ServiceNow interfaces
- A 2-3 second delay is used between operations to ensure page elements load properly

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Disclaimer
This project is not affiliated with or endorsed by ServiceNow. Use these automation scripts responsibly and in accordance with your organization's policies for system interaction.