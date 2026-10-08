# Energix Communication Automation Suite

A Python desktop application for automating repetitive communication workflows from Excel-based recipient data.

## Portfolio Edition
This public repository is sanitized. It contains no real recipient lists, phone numbers, email addresses, credentials, WhatsApp session data, or private documents.

## Core Features

### WhatsApp Automation
- Bulk messaging from Excel
- Phone-number normalization
- Multi-line message support
- WhatsApp Web automation with Selenium
- Optional PDF attachments
- Reusable message templates

### Email Automation
- Bulk personalized email sending
- Excel-driven recipient fields
- Custom subject and message templates
- Optional certificate attachment
- SMTP-based delivery
- Email app passwords are entered at runtime and are never persisted by the public build

### Certificate Automation
- Generate certificates from an image template and Excel names
- Arabic/RTL name support
- Preview before batch generation
- Adjustable font size and name position
- PDF output

### Enhanced WhatsApp Prototype
The second source file demonstrates a newer WhatsApp workflow with Excel/CSV support, configurable delay, optional persistent Chrome profile, per-recipient/default PDF paths, and a runtime log panel.

## Tech Stack
Python • Tkinter • Pandas • Selenium • Pillow • SMTP • OpenPyXL • Requests • Arabic Reshaper • Python Bidi

## Run
```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python src/energix_suite.py
```

## Privacy
Use only synthetic/demo data in this public repository. Never commit SMTP passwords, WhatsApp/Chrome profiles, real spreadsheets, generated certificates, or customer/student data.
