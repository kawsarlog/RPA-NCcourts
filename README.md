# RPA-NCcourts
🐍📜 This repository contains a Python script that utilizes the official API🔍 of **RPA WebService** to extract data and save it into an Excel file. The script uses various Python modules for automation and data manipulation💻


# RPA WebService Data Extractor

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue)](https://www.python.org/downloads/)
[![Selenium](https://img.shields.io/badge/Selenium-4.13.0-brightgreen)](https://pypi.org/project/selenium/)
[![Webdriver Manager](https://img.shields.io/badge/Webdriver%20Manager-4.0.0-brightgreen)](https://pypi.org/project/webdriver-manager/)
[![Openpyxl](https://img.shields.io/badge/Openpyxl-3.1.2-brightgreen)](https://pypi.org/project/openpyxl/)
[![Requests](https://img.shields.io/badge/Requests-2.31.0-brightgreen)](https://pypi.org/project/requests/)
[![Beautiful Soup](https://img.shields.io/badge/Beautiful%20Soup-4-brightgreen)](https://pypi.org/project/beautifulsoup4/)

This repository contains a Python script that utilizes the official API of **RPA WebService** to extract data and save it into a CSV file. The script uses various Python modules for automation and data manipulation.

## Prerequisites

Before you can use this script, ensure you have the following dependencies installed:

- **Python 3.9 or higher**: Make sure you have Python 3 installed on your system.
- **Required Python Modules**:
  - `selenium==4.13.0`
  - `webdriver-manager==4.0.0`
  - `openpyxl==3.1.2`
  - `requests==2.31.0`
  - `beautifulsoup4`

Update the following lines with your **RPA WebService** credentials:

   ```python
   email = 'your_email@example.com'
   password = 'your_password'
   ```

Inspect `category_code.txt` which contains all the necessary category codes for `case_type` and `partyExtendedConnectionTypes`.
