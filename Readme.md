# Partime Scheduler → Google Calendar Importer

## Overview

This project is a **Python automation script** that automatically reads your **Part-time weekly schedule email**, extracts the **Excel attachment**, parses your work shifts, and adds them to **Google Calendar**.

The script runs locally and can be scheduled using **cron** so your calendar stays automatically updated.

### Workflow

```
Gmail → Download Excel attachment → Parse shifts → Add events to Google Calendar
```

The script also keeps a small local state file to prevent duplicate imports.

---

# Features

* Automatically scans Gmail for schedule emails
* Downloads `.xls` or `.xlsx` attachments
* Converts old `.xls` files to `.xlsx`
* Extracts your shifts from the roster
* Adds shifts to Google Calendar
* Prevents duplicate calendar events
* Designed to run automatically with cron

---

# Project Structure

```
Python_calender/
│
├── app/
│   └── main.py
│
├── data/
│   ├── attachments/
│   ├── state.json
│   └── token.json
│
├── credentials.json
├── run_import.sh
└── README.md
```

### File Descriptions

| File               | Purpose                                                  |
| ------------------ | -------------------------------------------------------- |
| `main.py`          | Main script that reads email and creates calendar events |
| `credentials.json` | OAuth credentials from Google Cloud                      |
| `token.json`       | Stored authentication token                              |
| `state.json`       | Keeps track of processed emails and events               |
| `attachments/`     | Temporary storage for downloaded Excel files             |
| `run_import.sh`    | Script executed by cron                                  |

---

# Requirements

* Python **3.9+**
* Gmail account
* Google Calendar
* LibreOffice (for `.xls` conversion)

---

# Installation

## 1. Clone the repository

```
git clone <your_repo>
cd Python_calender
```

---

## 2. Create virtual environment

```
python3 -m venv venv
source venv/bin/activate
```

---

## 3. Install dependencies

```
pip install google-api-python-client
pip install google-auth
pip install google-auth-oauthlib
pip install openpyxl
pip install python-dateutil
```

---

## 4. Install LibreOffice (required for `.xls` files)

```
brew install --cask libreoffice
```

---

# Google API Setup

## 1. Go to Google Cloud Console

Create a new project.

---

## 2. Enable APIs

Enable:

* Gmail API
* Google Calendar API

---

3. Create OAuth credentials

Create OAuth Client ID

Type:

Desktop Application

Download the file and save it as:

credentials.json

in the project root.

First Run (Authentication)

Run the script manually once:

export DATA_DIR="$PWD/data"
export GOOGLE_CREDS_JSON="$PWD/credentials.json"
export YOUR_NAME="Your Name In Schedule"
export CALENDAR_ID="primary"
export GMAIL_QUERY='has:attachment (filename:xlsx OR filename:xls) newer_than:60d'

python3 app/main.py

A browser window will open asking for permission.

After authentication the script will create:

data/token.json

Future runs will not require login.

Gmail Search Query

Example query:

has:attachment (filename:xlsx OR filename:xls) newer_than:60d

You can improve it by filtering the sender:

Example:

from:primark has:attachment filename:xls
Cron Automation

Create a wrapper script.

run_import.sh
#!/bin/bash

cd ~/GitHub/Python_calender
source venv/bin/activate

export DATA_DIR="$PWD/data"
export GOOGLE_CREDS_JSON="$PWD/credentials.json"
export YOUR_NAME="Your Name In Schedule"
export CALENDAR_ID="primary"
export GMAIL_QUERY='has:attachment (filename:xlsx OR filename:xls) newer_than:60d'

python3 app/main.py >> "$PWD/data/cron.log" 2>&1

Make it executable:

chmod +x run_import.sh
Create Cron Job

Open crontab:

crontab -e

Example schedule (every 4 hours):

0 */4 * * * /bin/bash /Users/likhithkumara/GitHub/Python_calender/run_import.sh
Logs

Cron logs are saved to:

data/cron.log
State Management

The script stores state in:

data/state.json

This file tracks:

processed Gmail messages

created calendar events

This prevents duplicate events if cron runs multiple times.

Example Calendar Event
Title: Penneys / Primark Shift
Time: 10:00 – 14:30
Location: (optional)
Troubleshooting
Gmail API Error

Enable Gmail API in Google Cloud.

File contains no valid workbook part

Occurs when .xls files are opened using openpyxl.
The script converts them using LibreOffice.

Name not found

Ensure the name matches the roster exactly.

Example:

Likhith Kumar Arun Kumar
Future Improvements

Possible improvements:

Auto-detect your name from Gmail address

Create separate calendar for work shifts

Send Slack/Discord notification when schedule arrives

Detect shift updates and modify calendar events

Support multiple employees

Package as a CLI tool

Author

Likhith Kumar Arun Kumar

Project built to automate schedule management for Primark/Penneys shifts.
