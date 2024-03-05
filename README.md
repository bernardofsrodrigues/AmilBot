# AmilBot

# Web Page Query Automation using Selenium

This is a Python script that automates the process of querying a web page using Selenium WebDriver. The script reads data from an Excel file, fills out forms on a web page, performs queries, and extracts relevant information.

## Requirements

- Python 3.x
- Python libraries: selenium, pandas, openpyxl, PySimpleGUI

## Installation

1. Clone this repository:

    ```
    git clone https://github.com/bernardofsrodrigues/AmilBot.git
    ```

2. Install dependencies:

    ```
    pip install -r requirements.txt
    ```

## Usage

1. Run the `main.py` script.
2. Fill out the form with your username, password, and select an Excel file containing the data required for the query.
3. Click the "Submit" button to start the query process.
4. Wait until the script finishes processing the data and extracting the information.

## Features

- Reading data from an Excel file.
- Authentication on a web page.
- Automatic form filling.
- Performing queries and extracting information.

## Notes

- Make sure to provide correct authentication information to access the web page.
- The script has been tested on a Windows environment and may require adjustments to work on other operating systems.
