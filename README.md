# Absent Clients Report Automation

[![Python application](https://github.com/marcosirenos/Absent-Clients/actions/workflows/python-app.yml/badge.svg)](https://github.com/marcosirenos/Absent-Clients/actions/workflows/python-app.yml)

This project automates the entire workflow of generating an "Absent Clients" report. It handles data extraction from a web tool, processes the downloaded data, and generates a final, clean Excel report.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Workflow](#workflow)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
  - [Configuration](#configuration)
- [Usage](#usage)
- [Roadmap](#roadmap)
- [Continuous Integration](#continuous-integration)

---

## Overview

The primary objective of this project is to eliminate the manual effort involved in creating the Absent Clients report. The process involves logging into the Flex Monitor tool, applying specific filters, downloading a raw `.xls` file, cleaning and transforming the data based on business logic, and producing a multi-sheet `.xlsx` file ready for analysis.

## Features

*   **Automated Web Scraping**: Uses Selenium to navigate the Flex Monitor web tool, handle iFrames, and download the necessary report.
*   **Robust Data Processing**:
    *   Automatically converts the downloaded `.xls` (HTML in disguise) file to a usable format without manual intervention.
    *   Applies a comprehensive set of business rules to clean, transform, and enrich the data using `pandas`.
    *   Connects to Google Sheets to fetch supplementary data for enrichment.
*   **Structured Output**: Generates a multi-sheet `.xlsx` file with pivoted, aggregated data for different regions (Paraná, Curitiba, etc.) and a sheet with the raw, processed data.

## Workflow

1.  **Data Extraction**: A Selenium-powered script logs into the Flex Monitor tool, navigates to the correct report, applies the required filters, and downloads the data as an `.xls` file.
2.  **File Preparation**: The downloaded `.xls` file, which is actually an HTML file, is read by `prepare_file.py`. It extracts the main data table and saves it as a clean `.xlsx` file in the `data/raw/` directory.
3.  **Data Transformation**: The core script `process_data.py` takes over. It reads the raw data, connects to Google Sheets to pull in coverage and market data, and then performs a series of transformations:
    *   Cleans and standardizes columns (e.g., 'Emissora TV', 'Anunciante').
    *   Filters out irrelevant advertisers (e.g., internal campaigns, political ads).
    *   Corrects and standardizes client location data based on a set of predefined rules.
    *   Calculates new metrics and determines market segments ('LOCAL', 'IMPORT', 'PREF', etc.).
4.  **Report Generation**: The script creates several pivot tables from the processed data, each tailored to a specific geographic region. These pivots, along with the cleaned base data, are saved into a single `.xlsx` file in the `data/processed/` directory.

## Getting Started

Follow these instructions to get a local copy up and running.

### Prerequisites

*   Python 3.10
*   A virtual environment is highly recommended.

### Installation

1.  Clone the repository:
    ```sh
    git clone https://github.com/Guilherme-Vso/Absent-Clients.git
    ```
2.  Navigate to the project directory:
    ```sh
    cd Absent-Clients
    ```
3.  Create and activate a virtual environment:
    ```sh
    # Windows
    python -m venv venv
    .\venv\Scripts\activate
    
    # macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```
4.  Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

### Configuration

This project requires access to Google Sheets. You must configure a Google Cloud service account with permissions for the Google Sheets and Google Drive APIs.

1.  Follow the Google Cloud documentation to create a service account and download its JSON key file.
2.  Rename the downloaded JSON file to `credentials.json`.
3.  Place the `credentials.json` file inside the `dataprocessing/` directory. The path should be `dataprocessing/credentials.json`.
    > **Note**: This file is included in `.gitignore` and should never be committed to version control.
4.  Share your target Google Sheets with the service account's email address (found in the `client_email` field of your `credentials.json`).

The structure should look like this:
```
Absent-Clients/
└── dataprocessing/
    ├── credentials.json
    ├── credentails_example.json
    └── process_data.py
```

## Usage

To run the entire automated workflow, execute the main script from the root of the project directory:

```sh
python -m main.main
```

The final report will be saved in the `data/processed/` directory with a name corresponding to the current month (e.g., `July_.xlsx`).

## Roadmap

*   [ ] **Database Integration**: Replace the final Excel output with a direct data upload to a relational database (e.g., PostgreSQL, MySQL).
*   [ ] **Improved Logging**: Transition from `print()` statements to a structured logging library (e.g., `logging`) for better traceability and debugging.
*   [ ] **Error Handling & Retries**: Implement more robust error handling, especially for the web scraping part, with automatic retries on failure.
*   [ ] **Configuration File**: Move hardcoded values (like client-specific rules) into a separate configuration file (e.g., `config.ini` or `config.yaml`).

## Continuous Integration

This project uses GitHub Actions for Continuous Integration. The workflow, defined in `.github/workflows/python-app.yml`, automatically runs on every push and pull request to the `main` branch. It performs the following checks:

*   **Linting**: Uses `flake8` to enforce code style and check for syntax errors.
*   **Testing**: Runs the test suite using `pytest` to ensure core functionality remains stable.