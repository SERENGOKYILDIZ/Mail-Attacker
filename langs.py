# langs.py - Multi-language support for MailFlow
import locale

TRANSLATIONS = {
    "en": {
        # Sidebar
        "contacts": "👥 Contacts",
        "help": "❓ Help",
        "settings": "⚙️ Settings",
        "reports": "📊 Reports",
        "favorites": "⭐️ Favorites",
        "send_all_now": "🚀 Send All Now",

        # Contacts View
        "add_new_contact": "+ Add New Contact",
        "import_csv": "Import CSV",
        "export_csv": "Export CSV",
        "select_deselect_all": "Select/Deselect All",
        "search_placeholder": "Search contacts by company or email...",
        "sort_all": "All",
        "sort_active": "Active",
        "sort_inactive": "Inactive",
        "sort_favs": "Favorites",
        "sort_name": "By Name",
        "ready": "Ready",
        "edit": "Edit",
        "del": "Del",
        "moving_contacts": "Moving {n} Contact(s)...",

        # Context Menu
        "toggle_selected": "Toggle Selected",
        "toggle_favorite_selected": "Toggle Favorite (Selected)",
        "delete_selected": "Delete Selected",
        "select_all": "Select All",
        "deselect_all": "Deselect All",

        # Settings
        "app_settings": "Application Settings",
        "save_settings": "Save Settings",
        "smtp_config": "SMTP Configuration",
        "smtp_server_label": "SMTP Server Account (e.g. smtp.gmail.com)",
        "smtp_port_label": "SMTP Port (e.g. 587)",
        "sender_email": "Sender Email Address",
        "sender_password": "Sender Password / App Password",
        "other_settings": "Other Settings",
        "delay_label": "Delay Between Emails (Seconds)",
        "date_format": "Date Format",
        "time_format": "Time Format",
        "enable_notif": "Enable Desktop Notifications on Completion",
        "language": "Language",
        "data_folder": "Data Folder",
        "data_folder_desc": "Custom location for app data (config, reports, attachments). Leave empty for default (AppData).",
        "browse": "Browse",
        "reset_default": "Reset",
        "data_folder_moved": "Data folder changed and files have been moved. The app will restart.",
        "settings_saved": "Settings saved.",
        "lang_restart": "Language changed. Please restart the application for it to take full effect.",
        "theme": "Appearance Theme",
        "theme_system": "System",
        "theme_light": "Light",
        "theme_dark": "Dark",
        "global_signature": "Global Signature",
        "global_signature_desc": "Appended to messages using {signature}. Leave empty to disable.",

        # Reports
        "send_reports": "Send Reports",
        "clear_reports": "Clear Reports",
        "recent_logs": "Recent Logs",
        "clear_reports_confirm": "Are you sure you want to clear all report logs?",
        "send_batch": "Send Batch - {date} ({count} emails)",
        "delete_batch_confirm": "Are you sure you want to delete the report batch from {date}?",
        "message_details": "Message Details",
        "report_to": "To",
        "report_subject": "Subject",
        "no_message": "No message content recorded.",
        "attachments_label": "Attachments:",

        # Help
        "help_title": "MailFlow Help & Guide",
        "help_1_title": "1. Dynamic Template Variables",
        "help_1_body": (
            "You can personalize bulk emails easily by injecting variables directly into the 'Message Body' or 'Subject' field of any contact:\n\n"
            "• {company_name} : Replaced by the exact content of the Company Name field.\n"
            "• {email} : Replaced by the recipient's full email address.\n"
            "• {email_prefix} : Replaced by the first part of the email address (the text before the '@' symbol).\n"
            "• {date} : Replaced with today's date (e.g. 2026-03-08).\n"
            "• {time} : Replaced with the current time (e.g. 14:30).\n"
            "• {signature} : Replaced with the text set in Settings -> Global Signature.\n\n"
            "Example Message:\n'Hello {company_name} rep. We are contacting you at {time} via {email_prefix}.'\n"
            "Becomes: 'Hello Microsoft rep. We are contacting you at 14:30 via billgates.'"
        ),
        "help_2_title": "2. CSV Import / Export formatting",
        "help_2_body": (
            "When importing a .csv file, MailFlow looks at the header row to determine where data goes. To ensure a seamless import, please have your column headers loosely match these terms (case-insensitive):\n\n"
            "• Emails: 'email', 'e-mail', or 'mail'\n"
            "• Company: 'company', 'name', 'client', or 'contact'\n"
            "• Subject: 'subject' or 'title'\n"
            "• Message: 'message', 'body', 'text', or 'content'\n\n"
            "Note: Only rows with a valid email address column are imported. Attachments must be added manually inside the app afterwards."
        ),
        "help_3_title": "3. Anti-Spam & Send Rate",
        "help_3_body": (
            "To avoid your emails being flagged as spam or bot-activity by providers like Gmail or Office365, MailFlow gives you the ability to set a 'Delay Between Emails' in the Settings tab.\n"
            "A standard delay is 2 seconds, meaning MailFlow waits 2 seconds before firing off the next email in the queue. You can increase this delay for massive mailing campaigns to assure maximum deliverability."
        ),
        "help_4_title": "4. Desktop Notifications",
        "help_4_body": "Desktop notifications allow MailFlow to silently send emails in the background and pop-up a Windows toast message only when the campaign has entirely finished. You can explicitly disable this in the Settings tab.",

        # Add/Edit Contact Popup
        "add_new_contact_title": "Add New Contact",
        "edit_contact_title": "Edit Contact",
        "company_name": "Company Name",
        "recipient_email": "Recipient Email",
        "subject": "Subject",
        "message_body": "Message Body",
        "insert_label": "Insert:",
        "attachments": "Attachments",
        "no_files_selected": "No files selected",
        "add_select_files": "📎 Add Select File(s)",
        "open_folder": "📂 Open Folder",
        "save_contact_list": "Save Contact List",
        "add_to_contact_list": "Add to Contact List",
        "copy": "Copy",
        "paste": "Paste",
        "copying_files": "Copying files... ({current}/{total})",
        "email_required": "Recipient Email is required",
        "tag": "Category Tag",
        "all_tags": "All Tags",
        
        # Favorites
        "notes": "Notes",
        "save_note": "Save Note",
        "no_favorite_selected": "Select a favorite contact to view notes.",
        "note_saved_success": "Notes saved successfully.",

        # CSV
        "csv_import_success": "{count} contacts imported successfully.",
        "csv_no_email": "No valid rows found. Make sure a column header contains 'email'.",
        "csv_export_success": "{count} contacts exported successfully.",

        # Mail Engine
        "smtp_error": "Please configure SMTP settings first.",
        "sending": "Sending...",
        "sent_ok": "Sent ✓",
        "error_prefix": "Error",
        "skipped_disabled": "Skipped (Disabled)",
        "notif_title": "MailFlow - Sending Complete",
        "notif_body": "Sent: {sent} | Errors: {errors} | Skipped: {skipped}",

        # Dialogs
        "delete_confirm_title": "Confirm Deletion",
        "delete_confirm_msg": "Are you sure you want to delete the contact for '{email}'?",
        "delete_selected_confirm": "Are you sure you want to delete {count} selected contacts?",
        "invalid_email": "Please enter a valid email address.",
        "success": "Success",
        "error": "Error",
        "info": "Info",
    },
    "tr": {
        # Sidebar
        "contacts": "👥 Kişiler",
        "help": "❓ Yardım",
        "settings": "⚙️ Ayarlar",
        "reports": "📊 Raporlar",
        "favorites": "⭐️ Favoriler",
        "send_all_now": "🚀 Tümünü Gönder",

        # Contacts View
        "add_new_contact": "+ Yeni Kişi Ekle",
        "import_csv": "CSV İçe Aktar",
        "export_csv": "CSV Dışa Aktar",
        "select_deselect_all": "Tümünü Seç/Bırak",
        "search_placeholder": "Şirket veya e-posta ile kişi ara...",
        "sort_all": "Tümü",
        "sort_active": "Etkin",
        "sort_inactive": "Devre Dışı",
        "sort_favs": "Favoriler",
        "sort_name": "İsme Göre",
        "ready": "Hazır",
        "edit": "Düzenle",
        "del": "Sil",
        "moving_contacts": "{n} Kişi Taşınıyor...",

        # Context Menu
        "toggle_selected": "Seçilenleri Değiştir",
        "toggle_favorite_selected": "Seçilenleri Favori Yap/Kaldır",
        "delete_selected": "Seçilenleri Sil",
        "select_all": "Tümünü Seç",
        "deselect_all": "Tümünü Bırak",

        # Settings
        "app_settings": "Uygulama Ayarları",
        "save_settings": "Ayarları Kaydet",
        "smtp_config": "SMTP Yapılandırması",
        "smtp_server_label": "SMTP Sunucu Adresi (örn. smtp.gmail.com)",
        "smtp_port_label": "SMTP Port (örn. 587)",
        "sender_email": "Gönderici E-posta Adresi",
        "sender_password": "Gönderici Şifre / Uygulama Şifresi",
        "other_settings": "Diğer Ayarlar",
        "delay_label": "E-postalar Arası Bekleme (Saniye)",
        "date_format": "Tarih Formatı",
        "time_format": "Saat Formatı",
        "enable_notif": "Tamamlandığında Masaüstü Bildirimi Gönder",
        "language": "Dil",
        "data_folder": "Veri Klasörü",
        "data_folder_desc": "Uygulama verileri için özel konum (ayarlar, raporlar, ekler). Varsayılan (AppData) için boş bırakın.",
        "browse": "Gözat",
        "reset_default": "Sıfırla",
        "data_folder_moved": "Veri klasörü değiştirildi ve dosyalar taşındı. Uygulama yeniden başlatılacak.",
        "settings_saved": "Ayarlar kaydedildi.",
        "lang_restart": "Dil değiştirildi. Tam olarak uygulanması için uygulamayı yeniden başlatın.",
        "theme": "Görünüm Teması",
        "theme_system": "Sistem",
        "theme_light": "Aydınlık",
        "theme_dark": "Karanlık",
        "global_signature": "Genel İmza",
        "global_signature_desc": "Mesajlara {signature} kullanılarak eklenir. Kapatmak için boş bırakınız.",

        # Reports
        "send_reports": "Gönderim Raporları",
        "clear_reports": "Raporları Temizle",
        "recent_logs": "Son Kayıtlar",
        "clear_reports_confirm": "Tüm rapor kayıtlarını silmek istediğinize emin misiniz?",
        "send_batch": "Gönderim - {date} ({count} e-posta)",
        "delete_batch_confirm": "{date} tarihli rapor grubunu silmek istediğinize emin misiniz?",
        "message_details": "Mesaj Detayları",
        "report_to": "Alıcı",
        "report_subject": "Konu",
        "no_message": "Kayıtlı mesaj içeriği yok.",
        "attachments_label": "Ekler:",

        # Help
        "help_title": "MailFlow Yardım & Rehber",
        "help_1_title": "1. Dinamik Şablon Değişkenleri",
        "help_1_body": (
            "Toplu e-postalarınızı kişiselleştirmek için 'Mesaj Gövdesi' veya 'Konu' alanına doğrudan değişken ekleyebilirsiniz:\n\n"
            "• {company_name} : Şirket Adı alanının tam içeriği ile değiştirilir.\n"
            "• {email} : Alıcının tam e-posta adresi ile değiştirilir.\n"
            "• {email_prefix} : E-posta adresinin '@' işaretinden önceki kısmı ile değiştirilir.\n"
            "• {date} : Bugünün tarihi ile değiştirilir (örn. 2026-03-08).\n"
            "• {time} : Anlık saat ile değiştirilir (örn. 14:30).\n"
            "• {signature} : Ayarlar -> Genel İmza alanında yazdığınız metin ile değiştirilir.\n\n"
            "Örnek Mesaj:\n'Merhaba {company_name} yetkilisi. {time} itibarıyla {email_prefix} hesabınıza ulaşıyoruz.'\n"
            "Sonuç: 'Merhaba Microsoft yetkilisi. 14:30 itibarıyla billgates hesabınıza ulaşıyoruz.'"
        ),
        "help_2_title": "2. CSV İçe/Dışa Aktarma Formatı",
        "help_2_body": (
            "Bir .csv dosyasını içe aktarırken MailFlow, verilerin nereye gideceğini belirlemek için başlık satırına bakar. Sorunsuz aktarım için sütun başlıklarınız şu terimlere yakın olmalıdır (büyük/küçük harf duyarsız):\n\n"
            "• E-posta: 'email', 'e-mail' veya 'mail'\n"
            "• Şirket: 'company', 'name', 'client' veya 'contact'\n"
            "• Konu: 'subject' veya 'title'\n"
            "• Mesaj: 'message', 'body', 'text' veya 'content'\n\n"
            "Not: Yalnızca geçerli e-posta sütunu olan satırlar içe aktarılır. Ekler uygulama içinden manuel eklenmelidir."
        ),
        "help_3_title": "3. Anti-Spam & Gönderim Hızı",
        "help_3_body": (
            "E-postalarınızın Gmail veya Office365 gibi sağlayıcılar tarafından spam olarak işaretlenmesini önlemek için Ayarlar sekmesindeki 'E-postalar Arası Bekleme' özelliğini kullanabilirsiniz.\n"
            "Standart bekleme süresi 2 saniyedir. Büyük gönderim kampanyalarında bu süreyi artırarak maksimum teslim edilebilirlik sağlayabilirsiniz."
        ),
        "help_4_title": "4. Masaüstü Bildirimleri",
        "help_4_body": "Masaüstü bildirimleri, MailFlow'un e-postaları arka planda sessizce göndermesine ve kampanya tamamen bittiğinde size bir Windows bildirimi göstermesine olanak tanır. Bunu Ayarlar sekmesinden devre dışı bırakabilirsiniz.",

        # Add/Edit Contact Popup
        "add_new_contact_title": "Yeni Kişi Ekle",
        "edit_contact_title": "Kişiyi Düzenle",
        "company_name": "Şirket Adı",
        "recipient_email": "Alıcı E-posta",
        "subject": "Konu",
        "message_body": "Mesaj Gövdesi",
        "insert_label": "Ekle:",
        "attachments": "Ekler",
        "no_files_selected": "Dosya seçilmedi",
        "add_select_files": "📎 Dosya Ekle",
        "open_folder": "📂 Klasörü Aç",
        "save_contact_list": "Kişi Listesini Kaydet",
        "add_to_contact_list": "Kişi Listesine Ekle",
        "copy": "Kopyala",
        "paste": "Yapıştır",
        "copying_files": "Dosyalar kopyalanıyor... ({current}/{total})",
        "email_required": "Alıcı E-posta adresi gereklidir",
        "tag": "Kategori Etiketi",
        "all_tags": "Tüm Etiketler",
        
        # Favorites
        "notes": "Notlar",
        "save_note": "Notu Kaydet",
        "no_favorite_selected": "Notları görüntülemek için favori bir kişi seçin.",
        "note_saved_success": "Notlar başarıyla kaydedildi.",

        # CSV
        "csv_import_success": "{count} kişi başarıyla içe aktarıldı.",
        "csv_no_email": "Geçerli satır bulunamadı. Sütun başlığında 'email' olduğundan emin olun.",
        "csv_export_success": "{count} kişi başarıyla dışa aktarıldı.",

        # Mail Engine
        "smtp_error": "Lütfen önce SMTP ayarlarını yapılandırın.",
        "sending": "Gönderiliyor...",
        "sent_ok": "Gönderildi ✓",
        "error_prefix": "Hata",
        "skipped_disabled": "Atlandı (Devre Dışı)",
        "notif_title": "MailFlow - Gönderim Tamamlandı",
        "notif_body": "Gönderilen: {sent} | Hata: {errors} | Atlanan: {skipped}",

        # Dialogs
        "delete_confirm_title": "Silme Onayı",
        "delete_confirm_msg": "'{email}' kişisini silmek istediğinize emin misiniz?",
        "delete_selected_confirm": "Seçili {count} kişiyi silmek istediğinize emin misiniz?",
        "invalid_email": "Lütfen geçerli bir e-posta adresi girin.",
        "success": "Başarılı",
        "error": "Hata",
        "info": "Bilgi",
    }
}

_current_lang = "en"

def detect_system_language():
    """Detect system language and return 'tr' if Turkish, 'en' otherwise."""
    try:
        lang = locale.getdefaultlocale()[0]
        if lang and lang.startswith("tr"):
            return "tr"
    except Exception:
        pass
    return "en"

def set_language(lang_code):
    """Set the active language."""
    global _current_lang
    if lang_code in TRANSLATIONS:
        _current_lang = lang_code

def get_language():
    """Get the current language code."""
    return _current_lang

def t(key, **kwargs):
    """Translate a key using the current language. Supports format kwargs."""
    text = TRANSLATIONS.get(_current_lang, TRANSLATIONS["en"]).get(key, key)
    if kwargs:
        try:
            text = text.format(**kwargs)
        except (KeyError, IndexError):
            pass
    return text
