# Mass Email Sender

A Python script designed to automate sending personalized mass emails from an Excel contact list. It features rate-limiting, daily sending caps to prevent spam flagging, and progress tracking so it can be resumed safely if interrupted.

## Features
- Parses contact data from an Excel (`.xlsx`) file.
- Custom email templates with dynamic data injection (e.g., Name, Company).
- Automatically attaches a PDF resume/document.
- Built-in checkpointing (`sent_emails_log.txt`) ensures no duplicate emails are sent.
- Rate limits sending to 400 emails per session to stay within Gmail limits.
- Built-in delays to avoid triggering spam filters.

## Prerequisites
- Python 3.x installed.
- A Gmail account with an App Password generated. (Standard Google passwords won't work).

## Setup Instructions

### 1. Clone the repository
Make sure you are in the project folder.

### 2. Create a Virtual Environment (Recommended)
Creating a virtual environment ensures dependencies do not conflict with your system.
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install Requirements
Install the necessary Python libraries (`pandas`, `openpyxl`, `python-dotenv`):
```bash
pip install -r requirements.txt
```

### 4. Setup Configuration
1. Rename the `.env.example` file to `.env`.
2. Open the `.env` file and insert your Gmail address and 16-character App Password.
```env
SENDER_EMAIL=your_email@gmail.com
APP_PASSWORD=your_16_character_app_password
```

### 5. Add Your Files
Ensure the following files are in the same folder as the script:
1. **`contacts.xlsx`**: Your Excel sheet. Must contain columns with the names: `name`, `email`, and `company`. (Capitalization doesn't matter).
2. **`AnujMishraResume.pdf`**: Your resume or attachment. (If your resume has a different filename, change `RESUME_FILE_PATH` inside `send_mass_emails.py`).

## How to Run

To start the script, run the following command in your terminal:
```bash
source venv/bin/activate
python3 send_mass_emails.py
```

### Resuming the Script
The script caps out at 400 emails per session to protect your email account. 
If you have a list larger than 400, simply run the exact same command the next day. The script will automatically read `sent_emails_log.txt`, skip the people you already emailed, and instantly start sending to the next batch.

### Resetting the Campaign
If you want to start a brand new campaign using the same list, you must delete the log file first so it doesn't skip everyone:
```bash
rm sent_emails_log.txt
```
