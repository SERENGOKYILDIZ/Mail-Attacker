# MailFlow - Advanced Bulk Email Sender

## Project Overview

**MailFlow** is a powerful and modern Python-based desktop application designed for sending personalized bulk emails with targeted attachments. Built completely with CustomTkinter, it provides a sleek, dark-mode native interface that makes managing large contact lists, configuring SMTP settings, and tracking your email campaigns an effortless experience.

Whether you're sending out newsletters, corporate announcements, or individual tailored documents to a list of clients, MailFlow ensures that your emails are dispatched reliably while keeping your interface perfectly responsive.

## Downloads

[![Download MailFlow](https://img.shields.io/badge/Download-MailFlow_v1.2.0-blue?style=for-the-badge&logo=windows)](https://github.com/SERENGOKYILDIZ/Mail-Attacker/releases/download/v1.2.0/MailFlow_v1.2.0.exe)

## 👨‍💻 Author

**Author:** Semi Eren Gökyıldız
- **Email:** [gokyildizsemieren@gmail.com](mailto:gokyildizsemieren@gmail.com)
- **GitHub:** [SERENGOKYILDIZ](https://github.com/SERENGOKYILDIZ)
- **LinkedIn:** [Semi Eren Gökyıldız](https://www.linkedin.com/in/semi-eren-gokyildiz/)

---

## 🌟 Key Features

* **Modern, Responsive UI:** A beautiful dark-themed interface built with `customtkinter` featuring intuitive navigation between Contacts, Settings, and Reports.
* **Advanced Contact Management:** 
  * Add, edit, and delete contacts directly within a scrollable, proportionate grid view.
  * **Drag and Drop:** Easily reorder your mailing list by dragging contact cards.
  * **Multi-Selection:** Hold `CTRL` or `SHIFT` to select multiple contacts at once for bulk deletion or toggling.
  * **Enable/Disable:** Temporarily skip users during a bulk send without deleting them from your database.
* **Comprehensive Reporting System:** 
  * Every bulk send operation is logged chronologically under collapsible batch timestamps.
  * View exactly who received an email, what the subject was, and whether it was successfully delivered or encountered an error.
  * Click on any report row to view the exact body of the message and the attachment name sent to that user.
* **Bulletproof SMTP Configuration:** 
  * Offers preset configurations for major providers (Gmail, Office365, Yahoo) via a dropdown menu.
  * Built-in validation ensures ports are numeric and email formats are correct before saving.
  * Password visibility toggles for security and ease of use.
* **Threaded Mailing Engine & Dynamic Feedback:** 
  * The sending process runs entirely in the background. The UI remains fully interactive.
  * Contacts currently being processed are highlighted with a bright yellow border, providing real-time visual feedback of your campaign's progress.
* **Local Data Persistence:** 
  * All configurations, contact lists, reports, and attached files are automatically and securely cached in your local `AppData/Roaming/MailFlow` directory. 

## 🚀 Installation & Build

**Pre-built Executable:**
The easiest way to use MailFlow is to simply download the `.exe` from the Downloads section above. No Python installation required!

**From Source:**
1. Clone or download the repository.
2. Ensure you have **Python 3.11+** installed.
3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run the application:
   ```bash
   python app.py
   ```

## 📖 Usage Guide

1. **SMTP Setup:** Navigate to the **Settings** tab. Select your provider, enter your port (e.g., `587` for TLS), email, and password. 
   - *Note for Gmail/Office365 users:* You must generate and use an **App Password**; your regular account password will not work if 2-Step Verification is enabled.
2. **Contact Management:** Go to the **Contacts** tab to build your recipient list. You can assign unique subjects, tailored message bodies, and specific file attachments to every individual user.
3. **Sending:** Click **Send All** to initiate the bulk email process. Contacts will highlight in yellow as they are being processed.
4. **Reviewing Logs:** Head over to the **Reports** tab to review the historical success and details of your previous email campaigns.
