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

## 3. Create OAuth credentials

Create **OAuth Client ID**

Type:

```
Des
```
