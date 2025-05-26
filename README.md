# OWA Email Backup Tool

This Python script automates downloading and deleting emails from Outlook Web Access (OWA), saving each email as an `.eml` file directly from your mailbox.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
  - [Arguments](#arguments)
  - [Manual Steps During Execution](#manual-steps-during-execution)
- [Batch Script (Windows)](#batch-script-windows)
- [Important Notes](#important-notes)

## Features

- Automatically download emails from OWA as `.eml` files.
- Automatically delete emails after successful downloads.
- Supports accessing and backing up shared mailboxes.
- Browser automation via Microsoft Edge (Chromium) and Selenium.

## Requirements

- Python **3.10 or higher**
- Microsoft Edge browser
- Internet connection
- Python libraries:
  - `selenium`
  - `webdriver-manager`

## Installation

Clone this repository and install the required dependencies:

```bash
git clone https://github.com/Minimal-Project/OWA-Backup-Tool.git
cd OWA-Backup-Tool

pip install selenium webdriver-manager
```

---

## Usage

Execute the Python script directly from your terminal or command prompt:

```bash
python owa_downloader.py --output-dir "path/to/download_folder"
```

### Arguments

| Argument        | Default Value                           | Description                                        |
|-----------------|-----------------------------------------|----------------------------------------------------|
| `--owa-url`     | `https://outlook.office365.com/owa`     | URL to Outlook Web Access                          |
| `--output-dir`  | `downloads`                             | Directory to store downloaded `.eml` email files   |

#### Example with custom arguments:

```bash
python owa_downloader.py --owa-url "https://your-custom-owa-url" --output-dir "C:/EmailBackups"
```

### Manual Steps During Execution

The script guides you through a few manual steps:

1. **Login:**  
   Log into your Outlook Web Access mailbox in the browser that opens.

2. **Open Shared Mailbox (optional):**  
   If accessing a shared mailbox:
   - Click your profile icon.
   - Choose **"Open another mailbox."**
   - Enter the mailbox address and confirm.

3. **Prepare Mailbox:**  
   - Disable **conversation view** (`View → Show as → Messages`).
   - Navigate to the folder containing emails you wish to download.
   - Sort emails by date (**Oldest → Newest**).

After completing these steps, press **ENTER** in your terminal to start the automated download and deletion process.

---

## Batch Script (Windows)

A convenient Windows batch script (`run_backup.bat`) is included for quick execution:

### Usage

Simply double-click or run via Command Prompt:

```batch
run_backup.bat
```

### What this batch script does:

- Adds Python to your system PATH temporarily (if needed).
- Creates a `Downloads` folder in the script’s directory.
- Automatically launches the Python script to start the backup process.

## Important Notes

- **Data Backup:**  
  Ensure you securely back up your emails before running this script, as deleted emails are permanently removed.

- **Compatibility:**  
  This tool is specifically optimized for Microsoft Edge (Chromium-based).

- **Disclaimer:**  
  Use this tool responsibly and only with accounts you have explicit permission to access. The authors assume no liability for any misuse or data loss resulting from the use of this software.

---

© 2025 ReHoGa Interactive  

