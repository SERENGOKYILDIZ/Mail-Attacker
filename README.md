# MailFlow - Advanced Bulk Email Sender

## Project Overview

**MailFlow** is a powerful and modern Python-based desktop application designed for sending personalized bulk emails with targeted attachments. Built completely with CustomTkinter, it provides a sleek, dark-mode native interface that makes managing large contact lists, configuring SMTP settings, and tracking your email campaigns an effortless experience.

Whether you're sending out newsletters, corporate announcements, or individual tailored documents to a list of clients, MailFlow ensures that your emails are dispatched reliably while keeping your interface perfectly responsive.

## Downloads

[![Download MailFlow](https://img.shields.io/badge/Download-MailFlow_Latest-blue?style=for-the-badge&logo=windows)](https://github.com/SERENGOKYILDIZ/Mail-Attacker/releases/latest)

## 👨‍💻 Author

**Author:** Semi Eren Gökyıldız
- **Email:** [gokyildizsemieren@gmail.com](mailto:gokyildizsemieren@gmail.com)
- **GitHub:** [SERENGOKYILDIZ](https://github.com/SERENGOKYILDIZ)
- **LinkedIn:** [Semi Eren Gökyıldız](https://www.linkedin.com/in/semi-eren-gokyildiz/)

## 🌟 Key Features

* **Modern, Responsive UI:** A beautiful interface built with `customtkinter` featuring intuitive navigation between Contacts, Settings, Reports, and Help.
  * Choose between **Dark**, **Light**, or **System** appearance themes.
* **🌍 Multi-Language Support:**
  * Switch between **English** and **Türkçe** directly from Settings.
  * Auto-detects your system language on first launch.
  * Language change applies instantly — the app restarts automatically with the new language.
  * All UI strings, buttons, labels, dialogs, and help content are fully translated.
* **Advanced Contact Management:** 
  * Add, edit, and delete contacts directly within a scrollable, proportionate grid view.
  * **Optimized Two-Column Editor:** A two-column popup cleanly separates text inputs and attachments (left) from the message body (right).
  * **Favorites & Notes:** Star specific contacts as favorites and write persistent internal notes about them in a dedicated 'Favorites' sidebar panel.
  * **Categorization Tags (NEW):** Assign custom group text or tags to your contacts in the Editor to easily classify and organize your lists.
  * **Instant Sorting & Filtering:** Quickly filter your list using one-click buttons (All, Active, Inactive, Favorites), filter by specific **Tags**, or sort them alphabetically (A-Z) without any UI lag.
  * **Lightning Fast Navigation:** Enjoy accelerated 4x native mouse-wheel scrolling across all long lists and menus for a seamless browser-like experience.
  * **Clickable Template Variables:** Insert dynamic variables like `{company_name}`, `{email}`, `{email_prefix}`, `{date}`, and `{time}` into your message with a single click.
  * **Right-Click Context Menu:** Copy, paste, toggle, favorite, and delete contacts or message text via intuitive right-click menus.
  * **Multi-Attachment Support:** Attach multiple files to a single contact at once. Files copy over with an animated UI progress bar. Adding files preserves existing attachments — only new files are copied, and same-named files are silently updated.
  * **Dynamic Folder Sync:** Open your contact's attachment folder directly from the app. Any manual file changes you make in Windows Explorer automatically sync back to the app in real-time.
  * **CSV Import / Export:** Seamlessly migrate vast contact lists in seconds using standard `.csv` files.
  * **Live Search & Filter:** Quickly locate specific contacts using the real-time search bar.
  * **Multi-Selection:** Hold `CTRL`/`SHIFT` for bulk selection, toggle, favoriting, or deletion.
  * **Enable/Disable:** Temporarily skip users during a bulk send without deleting them from your database.
* **Smart Mailing Engine & Dynamic Feedback:** 
  * **Dynamic Template Variables:** Personalize emails at scale by inserting `{company_name}`, `{email}`, `{email_prefix}`, `{date}`, `{time}`, or the new `{signature}` directly into your subject or message body.
  * **Global Signature (NEW):** Set up a consistent, reusable email signature via the Settings tab to easily inject it into all outbound campaigns.
  * **Configurable Date & Time Formats:** Choose your preferred date format (DD/MM/YYYY, MM/DD/YYYY, YYYY-MM-DD, DD.MM.YYYY) and time format (24H / 12H) from read-only dropdown selectors in Settings.
  * **Anti-Spam Delay:** Bypass provider rate-limits by setting custom delays (in seconds) between each sent email. The delay is automatically skipped after the last enabled contact for faster completion.
  * The sending process runs entirely in the background. Contacts being processed are highlighted with a bright yellow border.
  * **Contextual Send All Button:** The prominent green "Send All Now" button only appears when you're on the Contacts tab, keeping other views clean.
  * **Desktop Notifications:** Receive a native Windows toast notification summarizing your campaign results (Sent, Errors, Skipped) the moment it finishes.
* **📂 Custom Data Folder (NEW):**
  * Store all app data (config, reports, attachments) in a custom location instead of AppData.
  * Perfect for backing up data to an external or fixed drive.
  * Files are automatically migrated when changing the data folder.
  * Attachment paths auto-resolve to the current data location — no broken links after moving.
  * A small pointer file (`data_folder.txt`) in AppData tracks where your data lives.
* **In-App Help & Guide:**
  * A dedicated Help tab with scrollable, detailed sections covering template variables, CSV formatting, anti-spam configuration, and desktop notifications.
* **Comprehensive Reporting System:** 
  * Every bulk send operation is logged chronologically under collapsible batch timestamps. Easily delete old batches with a single click.
  * View exactly who received an email, what the subject was, and whether it was successfully delivered or encountered an error.
  * Click on any report row to view the exact body of the message and a vertically formatted list of all attached files sent to that user.
* **Bulletproof SMTP Configuration:** 
  * Offers preset configurations for major providers (Gmail, Office365, Yahoo) via a dropdown menu.
  * Built-in validation ensures ports are numeric and email formats are correct before saving.
  * **Compact Settings Layout:** Server + Port share a row, Email + Password share a row, and Language + Delay + Notifications are grouped together for a cleaner, more efficient interface.
  * Password visibility toggles for security and ease of use.
  * **Encrypted Storage (NEW):** Your SMTP configuration passwords are safely encrypted on-disk (`crypto.py`) before they are written to the configuration file.
* **Local Data Persistence:** 
  * All configurations, contact lists, reports, and attached files are automatically and securely cached in your local `AppData/Roaming/MailFlow` directory (or your custom data folder).

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
2. **Language:** Choose your preferred language (English / Türkçe) from the Settings tab. The app will restart automatically to apply the change.
3. **Data Folder:** Optionally set a custom data folder from Settings to store all app data on a specific drive for backup safety.
4. **Contact Management:** Go to the **Contacts** tab to build your recipient list. You can assign unique subjects, tailored message bodies, and specific file attachments to every individual user.
5. **Sending:** Click **Send All** to initiate the bulk email process. Contacts will highlight in yellow as they are being processed.
6. **Reviewing Logs:** Head over to the **Reports** tab to review the historical success and details of your previous email campaigns.
