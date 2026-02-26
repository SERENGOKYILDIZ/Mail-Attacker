# Mail Attacker - Bulk Email Sender

## Project Overview

Mail Attacker is a modern, Python-based desktop application designed for sending personalized bulk emails with attachments. Built with CustomTkinter, it provides a sleek, user-friendly interface to manage contacts, configure SMTP settings, and send customized emails automatically.

## Key Features

* **Modern UI with Dark Mode:** A sleek, responsive interface built with CustomTkinter.
* **Integrated Contact Management:** Add, edit, and delete contacts directly within the scrollable list view.
* **Local Attachment Backup:** Selected attachments are automatically copied to a local `attachments/` folder to ensure reliability and prevent missing files during batch sending.
* **Threaded Mailing Engine:** The email sending process runs on a separate thread, ensuring the UI remains perfectly responsive during bulk operations.
* **Configurable Constants:** Centralized application settings (colors, dimensions, etc.) via `config.py` for easy customization.

## Installation

1. Clone or download the repository.
2. Ensure you have Python installed, then install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python app.py
   ```

## Usage

1. **SMTP Setup:** Navigate to the "Settings" tab to configure your SMTP server (e.g., `smtp.gmail.com`), port (e.g., `587`), and sender credentials. 
   - *Note for Gmail users:* You will need to generate and use an **App Password** if you have 2-Step Verification enabled.
2. **Contact Management:** Go to the "Contacts" tab to add recipients. You can include recipient-specific subjects, messages, and attachments.
3. **Sending:** Click "Send All" to initiate the bulk email process. The status of each email will update in real-time.
