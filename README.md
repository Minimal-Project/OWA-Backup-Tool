# OWA Email Backup Tool

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue)](https://www.python.org/)
[![Browser](https://img.shields.io/badge/Browser-Edge-blue)](https://www.microsoft.com/edge)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)](#)

A Python script to download and delete emails from Outlook Web Access (OWA), saving them as `.eml` files. Supports both personal and shared mailboxes. Uses Selenium with Microsoft Edge (Chromium).

## Features

- Downloads emails as `.eml` files from Outlook Web Access
- Deletes emails after successful download
- Supports personal and shared mailboxes
- Automates browser interaction using Selenium
- Includes a ready-to-use Windows batch script

## Requirements

- Python 3.10 or higher
- Microsoft Edge (Chromium)
- Internet connection
- Python packages:

```
pip install selenium webdriver-manager
```

## Installation

```
git clone https://github.com/Minimal-Project/OWA-Backup-Tool.git
cd OWA-Backup-Tool
pip install selenium webdriver-manager
```

## Usage

### Python Script

```
python owa_downloader.py --output-dir "C:/EmailBackups" --mailbox own
```

### Arguments

| Argument        | Default                              | Description                                      |
|----------------|--------------------------------------|--------------------------------------------------|
| --owa-url      | https://outlook.office365.com/owa    | URL to Outlook Web Access                        |
| --output-dir   | downloads                            | Directory where `.eml` files will be saved       |
| --mailbox      | own                                  | Mailbox type: `own` or `shared`                  |

### Example

```
python owa_downloader.py --owa-url "https://owa.example.com" --output-dir "C:/Backups" --mailbox shared
```

## Manual Steps During Execution

1. Login to Outlook Web Access in the opened browser window.
2. If using a shared mailbox:
   - Click the profile icon
   - Select "Open another mailbox"
   - Enter the shared address and confirm
3. In the mailbox:
   - Disable Conversation View (View → Show as → Messages)
   - Navigate to the desired folder
   - Sort emails by Oldest → Newest
4. Press ENTER in the terminal to start the download and deletion process.

## Batch Script (Windows)

The repository includes a batch script `run_backup.bat` to simplify running the tool.

### Usage

Double-click the script or run it in Command Prompt:

```
run_backup.bat
```

### What it does

- Adds Python to the PATH (temporarily, if needed)
- Creates a `Downloads` folder if it doesn't exist
- Prompts for mailbox type selection (`own` or `shared`)
- Launches the Python script with appropriate options
- Handles successful or failed exits and shows messages

### Prompt Example

```
Choose mailbox type:
  1) Own mailbox
  2) Shared mailbox
Enter your choice (1 or 2):
```

## Important Notes

- Emails are permanently deleted after download. Use with caution.
- Manual login is required due to security restrictions (SSO, MFA).
- Only Microsoft Edge (Chromium) is supported.
- Pressing `CTRL+C` will cancel the script cleanly and close the browser.
- Use this script only with accounts you are authorized to access.

---

© 2025 ReHoGa Interactive
