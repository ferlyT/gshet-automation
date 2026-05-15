# Google Sheets Billing Automation

A Python script designed to automate the synchronization of SQL Server billing data to Google Sheets. It features a modular, object-oriented structure, professional dashboard styling, advanced conditional formatting for data visualization, and robust error handling.

## Table of Contents
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Setup](#setup)
  - [1. Clone the Repository](#1-clone-the-repository)
  - [2. Install Dependencies](#2-install-dependencies)
  - [3. Google Cloud Project & Service Account Setup](#3-google-cloud-project--service-account-setup)
  - [4. Configure Environment Variables](#4-configure-environment-variables)
- [Usage](#usage)
- [Configuration (`.env` variables)](#configuration-env-variables)
- [Customization](#customization)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## Features

This project provides the following key functionalities:

-   **Modular and Object-Oriented Design**: Code is structured into `Config`, `DatabaseClient`, `GoogleSheetsClient`, and `BillingSyncApp` classes for better maintainability and scalability.
-   **Multi-Sheet Synchronization**: Capable of syncing data from different SQL queries to multiple Google Sheets within the same spreadsheet.
-   **Professional Dashboard Styling**:
    -   Frozen header rows (first 3 rows).
    -   Dark header (Row 3) with bold white text.
    -   Zebra banding (alternating white and light blue rows) for data readability starting from Row 4.
    -   Dynamic column sizing and row heights.
    -   "Last Update" timestamp on Row 1 with pastel red font in cell A1.
    -   Empty separator row (Row 2) with white background.
-   **Aging Visualization for 'hari' Column**:
    -   **3-Point Gradient Color Scale**: Cells in the 'hari' column are colored from green (0 days) to yellow (4 days) to red (7+ days) to visually represent aging.
    -   **Auto-Bold**: Numbers greater than 7 in the 'hari' column are automatically bolded for immediate attention.
    -   **Center Alignment**: 'hari' column values are center-aligned.
-   **Conditional Row Highlighting**: Entire rows are highlighted with a pastel red background if the `StatusKirim` column contains `'BELUM DITERIMA GDG'`.
-   **Dynamic Dropdown List for 'Sales' Column**: Automatically creates a data validation dropdown list for the 'Sales' column based on unique values present in the data.
-   **Automatic Whitespace Trimming**: All string data fetched from SQL is automatically trimmed to remove leading/trailing whitespaces.
-   **Robust Error Handling & Retry Mechanism**: Implements a retry mechanism with exponential backoff for Google Sheets API calls to handle rate limits (HTTP 429 errors).
-   **Logging**: Detailed logging for monitoring the synchronization process, including fetching data, applying styles, and syncing rows.

## Prerequisites

Before you begin, ensure you have the following:

-   **Python 3.7+**: Installed on your system.
-   **pip**: Python package installer.
-   **SQL Server**: Access to the SQL Server database specified in your configuration.
-   **Google Cloud Project**: A Google Cloud Project with the Google Sheets API and Google Drive API enabled.
-   **Google Service Account**: A service account key file (`credentials.json`) with appropriate permissions to access and modify your Google Sheet.

## Setup

### 1. Clone the Repository

If you haven't already, clone this repository to your local machine:

```bash
git clone https://github.com/ferlyT/gshet-automation.git
cd gshet-automation
```

### 2. Install Dependencies

It's recommended to use a virtual environment.

```bash
python -m venv .venv
source .venv/Scripts/activate  # On Windows
# source .venv/bin/activate    # On macOS/Linux

pip install -r requirements.txt
# If requirements.txt is not present, install manually:
# pip install pyodbc pandas gspread google-auth sqlalchemy python-dotenv
```

### 3. Google Cloud Project & Service Account Setup

1.  Go to the Google Cloud Console.
2.  Create a new project or select an existing one.
3.  Enable the **Google Sheets API** and **Google Drive API** for your project.
4.  Create a new **Service Account**:
    -   Navigate to `IAM & Admin > Service Accounts`.
    -   Click `+ CREATE SERVICE ACCOUNT`.
    -   Give it a name and grant it the `Editor` role (or more specific roles like `Google Sheets Editor` and `Google Drive Editor`).
    -   After creation, click on the service account, go to `Keys`, and click `ADD KEY > Create new key > JSON`.
    -   Download the JSON file and save it as `credentials.json` in the root directory of this project.
5.  **Share your Google Sheet** with the email address of your service account (found in the `client_email` field of `credentials.json`).

### 4. Configure Environment Variables

Create a file named `.env` in the root directory of your project (next to `main.py`) and fill it with your database and Google Sheet details.

```env
# .env example
DB_DRIVER={ODBC Driver 17 for SQL Server}
DB_SERVER=your_sql_server_address
DB_NAME=your_database_name
DB_USER=your_db_username
DB_PASS=your_db_password

GOOGLE_SHEET_ID=your_google_sheet_id_or_url
GOOGLE_CREDENTIALS_FILE=credentials.json
```

**Note**: The `GOOGLE_SHEET_ID` can be the full URL of your Google Sheet or just the ID extracted from the URL (e.g., `1abc2def3ghi4jkl5mno6pqr7stu8vwx9yz`).

## Usage

Once everything is set up, you can run the script from your terminal:

```bash
python main.py
```

The script will fetch data from your SQL Server, process it, and update the specified Google Sheets with the applied styling and conditional formatting.

## Configuration (`.env` variables)

-   `DB_DRIVER`: ODBC driver for SQL Server. Default is `{ODBC Driver 17 for SQL Server}`.
-   `DB_SERVER`: Your SQL Server's address (e.g., `localhost`, `192.168.1.100`, or a named instance).
-   `DB_NAME`: The name of the database to connect to.
-   `DB_USER`: Username for database authentication.
-   `DB_PASS`: Password for database authentication.
-   `GOOGLE_SHEET_ID`: The ID or full URL of your target Google Spreadsheet.
-   `GOOGLE_CREDENTIALS_FILE`: Path to your service account JSON key file. Default is `credentials.json`.

## Customization

You can easily add more sheets to be synchronized by modifying the `tasks` list in the `main()` function:

```python
def main():
    tasks = [
        {
            "sheet_name": "Udara",
            "query": "EXEC get_data_billing_gsheet 1" # Your SQL query for 'Udara' sheet
        },
        {
            "sheet_name": "Laut",
            "query": "EXEC get_data_billing_gsheet 2" # Your SQL query for 'Laut' sheet
        },
        # Add more tasks here
        # {
        #     "sheet_name": "NewSheetName",
        #     "query": "SELECT * FROM YourNewTable"
        # }
    ]

    app = BillingSyncApp()
    app.run(tasks)
```

## Troubleshooting

-   **`AttributeError: 'DatabaseClient' object has no attribute 'fetch_query'`**: This indicates a mismatch in method names. Ensure `fetch_query` is correctly defined and called. (This has been fixed in the latest version).
-   **`gspread.exceptions.APIError: {'code': 429, ...}`**: You've hit Google's API rate limit. The script has a built-in retry mechanism with exponential backoff to handle this. If it persists, reduce the frequency of runs or contact Google Cloud support.
-   **`gspread.exceptions.WorksheetNotFound`**: The specified sheet name does not exist in your Google Spreadsheet. The script will attempt to create it automatically.
-   **`gspread.exceptions.SpreadsheetNotFound`**: The `GOOGLE_SHEET_ID` in your `.env` file is incorrect or the service account does not have access.
-   **Database Connection Errors**: Double-check your `DB_DRIVER`, `DB_SERVER`, `DB_NAME`, `DB_USER`, and `DB_PASS` in the `.env` file. Ensure your SQL Server is accessible from where the script is running.
-   **`credentials.json` not found or invalid**: Ensure the `credentials.json` file is in the correct path and has the right content. Also, ensure the service account email is shared with your Google Sheet.

## License

This project is open-source and available under the MIT License.
