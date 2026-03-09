import os
import sys
import json
import shutil
import threading
import smtplib
import csv
from email.message import EmailMessage
import mimetypes
import datetime
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox

import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("eren.mailattacker.v1")

import config
import langs
import crypto
from langs import t

ctk.set_appearance_mode(config.DEFAULT_THEME)
ctk.set_default_color_theme(config.ACCENT_COLOR)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def resolve_att_path(abs_path):
    """Resolve an attachment path. If it exists, return as-is.
    Otherwise, extract profile/filename and resolve against current ATTACHMENTS_DIR."""
    if not abs_path:
        return abs_path
    if os.path.exists(abs_path):
        return abs_path
    # Try to extract the relative part: attachments/<profile>/<filename>
    parts = abs_path.replace("\\", "/").split("/")
    try:
        att_idx = [p.lower() for p in parts].index("attachments")
        rel = os.path.join(*parts[att_idx + 1:])
        resolved = os.path.join(config.ATTACHMENTS_DIR, rel)
        if os.path.exists(resolved):
            return resolved
    except (ValueError, TypeError):
        pass
    return abs_path

class MailAttackerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(f"{config.APP_NAME} {config.APP_VERSION}")
        self.geometry(f"{config.WINDOW_WIDTH}x{config.WINDOW_HEIGHT}")
        self.resizable(config.RESIZABLE, config.RESIZABLE)
        
        icon_path = resource_path(os.path.join("assets", "logo.ico"))
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        
        self.app_config = {"smtp": {}, "contacts": []}
        self.contact_widgets = {}
        self.load_config()
        self.apply_theme_and_lang()

    def apply_theme_and_lang(self):
        # Apply theme
        saved_theme = self.app_config.get("theme", "System")
        ctk.set_appearance_mode(saved_theme)
        
        # Initialize language
        saved_lang = self.app_config.get("language", "")
        if saved_lang and saved_lang in langs.TRANSLATIONS:
            langs.set_language(saved_lang)
        else:
            langs.set_language(langs.detect_system_language())

        if not os.path.exists(config.ATTACHMENTS_DIR):
            os.makedirs(config.ATTACHMENTS_DIR)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text=f"{config.APP_NAME}\n{config.APP_VERSION}", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.contacts_button = ctk.CTkButton(self.sidebar_frame, text=t("contacts"), command=self.show_contacts_view)
        self.contacts_button.grid(row=1, column=0, padx=20, pady=10)

        self.help_button = ctk.CTkButton(self.sidebar_frame, text=t("help"), command=self.show_help_view, fg_color="#3B3B3B", hover_color="#2B2B2B")
        self.help_button.grid(row=2, column=0, padx=20, pady=10)
        
        self.favorites_button = ctk.CTkButton(self.sidebar_frame, text=t("favorites"), command=self.show_favorites_view, fg_color="#c29400", hover_color="#9e7800", text_color="white")
        self.favorites_button.grid(row=3, column=0, padx=20, pady=10)
        
        self.reports_button = ctk.CTkButton(self.sidebar_frame, text=t("reports"), command=self.show_reports_view)
        self.reports_button.grid(row=4, column=0, padx=20, pady=10)
        
        self.settings_button = ctk.CTkButton(self.sidebar_frame, text=t("settings"), command=self.show_settings_view)
        self.settings_button.grid(row=5, column=0, padx=20, pady=10)
        
        self.send_all_button = ctk.CTkButton(self.sidebar_frame, text=t("send_all_now"), command=self.send_all_mails, fg_color="green", hover_color="darkgreen", height=45, font=ctk.CTkFont(size=14, weight="bold"))
        self.send_all_button.grid(row=7, column=0, padx=15, pady=(15, 25))

        # --- Contacts Frame ---
        self.contacts_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        self.add_contact_frame = ctk.CTkFrame(self.contacts_frame, fg_color="transparent")
        self.add_contact_frame.pack(fill="x", padx=config.PAD_X, pady=(20, 5))
        
        self.add_btn_top = ctk.CTkButton(self.add_contact_frame, text=t("add_new_contact"), fg_color="#1f538d", command=lambda: self.open_contact_popup(None))
        self.add_btn_top.pack(side="left")
        
        self.import_btn = ctk.CTkButton(self.add_contact_frame, text=t("import_csv"), fg_color="#915200", hover_color="#633800", command=self.import_csv)
        self.import_btn.pack(side="left", padx=(10, 0))
        
        self.export_btn = ctk.CTkButton(self.add_contact_frame, text=t("export_csv"), fg_color="#186121", hover_color="#0d3b13", command=self.export_csv)
        self.export_btn.pack(side="left", padx=(10, 0))
        
        self.toggle_all_var = ctk.BooleanVar(value=True)
        self.toggle_all_btn = ctk.CTkButton(self.add_contact_frame, text=t("select_deselect_all"), fg_color="gray30", hover_color="gray40", command=self.toggle_all_contacts)
        self.toggle_all_btn.pack(side="right")
        
        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(self.contacts_frame, textvariable=self.search_var, placeholder_text=t("search_placeholder"))
        self.search_entry.pack(fill="x", padx=config.PAD_X, pady=(0, 5))
        self.search_entry.bind("<KeyRelease>", self.filter_contacts_list)

        # Sort and Filter Panel
        self.sort_frame = ctk.CTkFrame(self.contacts_frame, fg_color="transparent")
        self.sort_frame.pack(fill="x", padx=config.PAD_X, pady=(0, 5))

        self.sort_var = ctk.StringVar(value="all")
        
        self.btn_sort_all = ctk.CTkRadioButton(self.sort_frame, text=t("sort_all"), variable=self.sort_var, value="all", command=self.apply_sorting)
        self.btn_sort_all.pack(side="left", padx=(0, 10))
        
        self.btn_sort_active = ctk.CTkRadioButton(self.sort_frame, text=t("sort_active"), variable=self.sort_var, value="active", command=self.apply_sorting)
        self.btn_sort_active.pack(side="left", padx=(0, 10))
        
        self.btn_sort_inactive = ctk.CTkRadioButton(self.sort_frame, text=t("sort_inactive"), variable=self.sort_var, value="inactive", command=self.apply_sorting)
        self.btn_sort_inactive.pack(side="left", padx=(0, 10))

        self.btn_sort_favs = ctk.CTkRadioButton(self.sort_frame, text=t("sort_favs"), variable=self.sort_var, value="favs", command=self.apply_sorting)
        self.btn_sort_favs.pack(side="left", padx=(0, 10))

        self.tag_filter_var = ctk.StringVar(value=t("all_tags"))
        self.btn_filter_tag = ctk.CTkComboBox(self.sort_frame, variable=self.tag_filter_var, values=[t("all_tags")], command=self.apply_sorting, width=130, state="readonly")
        self.btn_filter_tag.pack(side="left", padx=(0, 10))

        # This one is a checkbox/button overlay since "By Name" is an ordering override, but we can treat it as part of the radio group or a separate toggle. 
        # Using a separate button makes more sense for A-Z sorting that stacks with other filters.
        self.sort_name_var = ctk.BooleanVar(value=False)
        self.btn_sort_name = ctk.CTkCheckBox(self.sort_frame, text=t("sort_name"), variable=self.sort_name_var, command=self.apply_sorting)
        self.btn_sort_name.pack(side="right")

        self.contacts_scrollable_frame = ctk.CTkScrollableFrame(self.contacts_frame)
        self.contacts_scrollable_frame.pack(fill="both", expand=True, padx=config.PAD_X, pady=(5, 20))
        
        # Context Menu for Right Click
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label=t("toggle_selected"), command=self.toggle_selected_contacts)
        self.context_menu.add_command(label=t("toggle_favorite_selected"), command=self.toggle_favorite_selected_contacts)
        self.context_menu.add_command(label=t("delete_selected"), command=self.delete_selected_contacts)
        self.context_menu.add_separator()
        self.context_menu.add_command(label=t("select_all"), command=lambda: self.set_all_contacts(True))
        self.context_menu.add_command(label=t("deselect_all"), command=lambda: self.set_all_contacts(False))
        
        self.contacts_scrollable_frame.bind("<Button-3>", self.show_context_menu)
        
        # --- Favorites Frame ---
        self.favorites_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        fav_main_frame = ctk.CTkFrame(self.favorites_frame, fg_color="transparent")
        fav_main_frame.pack(fill="both", expand=True, padx=config.PAD_X, pady=20)
        
        # Left side: list of favorites
        self.fav_list_frame = ctk.CTkScrollableFrame(fav_main_frame, width=250)
        self.fav_list_frame.pack(side="left", fill="y", padx=(0, 10))
        
        # Right side: notes area
        self.fav_notes_frame = ctk.CTkFrame(fav_main_frame)
        self.fav_notes_frame.pack(side="right", fill="both", expand=True)
        
        self.active_fav_contact_id = None
        self.fav_title_lbl = ctk.CTkLabel(self.fav_notes_frame, text=t("no_favorite_selected"), font=ctk.CTkFont(size=18, weight="bold"))
        self.fav_title_lbl.pack(pady=15)
        
        self.fav_notes_textbox = ctk.CTkTextbox(self.fav_notes_frame, font=ctk.CTkFont(size=14))
        self.fav_notes_textbox.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        self.fav_notes_textbox.configure(state="disabled")
        
        # Auto-save notes on typing
        def on_note_modified(event):
            if self.active_fav_contact_id is not None:
                for c in self.app_config["contacts"]:
                    if c.get("id") == self.active_fav_contact_id:
                        c["notes"] = self.fav_notes_textbox.get("1.0", "end-1c")
                        self.save_config()
                        break
                        
        self.fav_notes_textbox.bind("<KeyRelease>", on_note_modified)

        # --- Settings Frame ---
        self.settings_frame = ctk.CTkScrollableFrame(self, corner_radius=0, fg_color="transparent")
        settings_top_frame = ctk.CTkFrame(self.settings_frame, fg_color="transparent")
        settings_top_frame.pack(fill="x", padx=config.PAD_X, pady=20)
        
        self.settings_title = ctk.CTkLabel(settings_top_frame, text=t("app_settings"), font=ctk.CTkFont(size=20, weight="bold"))
        self.settings_title.pack(side="left")
        
        save_btn = ctk.CTkButton(settings_top_frame, text=t("save_settings"), command=self.save_settings, width=150, fg_color="#1f538d")
        save_btn.pack(side="right")
        
        self.smtp_server_var = ctk.StringVar(value=self.app_config["smtp"].get("server", ""))
        self.smtp_port_var = ctk.StringVar(value=self.app_config["smtp"].get("port", ""))
        self.smtp_user_var = ctk.StringVar(value=self.app_config["smtp"].get("user", ""))
        
        # Decrypt password for display in settings
        saved_pass = self.app_config["smtp"].get("password", "")
        self.smtp_pass_var = ctk.StringVar(value=crypto.decrypt(saved_pass))

        smtp_frame = ctk.CTkFrame(self.settings_frame)
        smtp_frame.pack(fill="x", padx=config.PAD_X, pady=(0, 10))
        
        ctk.CTkLabel(smtp_frame, text=t("smtp_config"), font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", padx=config.PAD_X, pady=(10, 0))
        
        # Server + Port (same row)
        smtp_row1 = ctk.CTkFrame(smtp_frame, fg_color="transparent")
        smtp_row1.pack(fill="x", padx=config.PAD_X, pady=(5, 0))
        
        server_col = ctk.CTkFrame(smtp_row1, fg_color="transparent")
        server_col.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkLabel(server_col, text=t("smtp_server_label"), anchor="w").pack(fill="x")
        server_options = ["smtp.gmail.com", "smtp.office365.com", "smtp.mail.yahoo.com"]
        self.smtp_server_cb = ctk.CTkComboBox(server_col, values=server_options, variable=self.smtp_server_var, state="readonly")
        self.smtp_server_cb.pack(fill="x", pady=(2, 0))
        
        vcmd_port = (self.register(self.validate_port), '%P')
        port_col = ctk.CTkFrame(smtp_row1, fg_color="transparent")
        port_col.pack(side="right")
        ctk.CTkLabel(port_col, text=t("smtp_port_label"), anchor="w").pack(fill="x")
        ctk.CTkEntry(port_col, textvariable=self.smtp_port_var, validate='key', validatecommand=vcmd_port, width=80).pack(pady=(2, 0))
        
        # Row 2: Email + Password
        smtp_row2 = ctk.CTkFrame(smtp_frame, fg_color="transparent")
        smtp_row2.pack(fill="x", padx=config.PAD_X, pady=(8, 10))
        
        email_col = ctk.CTkFrame(smtp_row2, fg_color="transparent")
        email_col.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkLabel(email_col, text=t("sender_email"), anchor="w").pack(fill="x")
        self.email_entry = ctk.CTkEntry(email_col, textvariable=self.smtp_user_var, placeholder_text="example@domain.com")
        self.email_entry.pack(fill="x", pady=(2, 0))
        self.email_entry.bind("<FocusOut>", self.validate_email_format)
        
        pass_col = ctk.CTkFrame(smtp_row2, fg_color="transparent")
        pass_col.pack(side="left", fill="x", expand=True, padx=(5, 0))
        ctk.CTkLabel(pass_col, text=t("sender_password"), anchor="w").pack(fill="x")
        pass_inner = ctk.CTkFrame(pass_col, fg_color="transparent")
        pass_inner.pack(fill="x", pady=(2, 0))
        self.pass_entry = ctk.CTkEntry(pass_inner, textvariable=self.smtp_pass_var, show="*")
        self.pass_entry.pack(side="left", fill="x", expand=True)
        self.show_pass_var = ctk.BooleanVar(value=False)
        self.show_pass_btn = ctk.CTkCheckBox(pass_inner, text="👁", variable=self.show_pass_var, command=self.toggle_password_visibility, width=30)
        self.show_pass_btn.pack(side="right", padx=(5, 0))
        
        others_frame = ctk.CTkFrame(self.settings_frame)
        others_frame.pack(fill="x", padx=config.PAD_X, pady=(5, 10))
        
        ctk.CTkLabel(others_frame, text=t("other_settings"), font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", padx=config.PAD_X, pady=(10, 0))
        
        # Row 3: Language + Delay + Notifications
        misc_row = ctk.CTkFrame(others_frame, fg_color="transparent")
        misc_row.pack(fill="x", padx=config.PAD_X, pady=(5, 8))
        
        lang_col = ctk.CTkFrame(misc_row, fg_color="transparent")
        lang_col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(lang_col, text=t("language"), anchor="w").pack(fill="x")
        lang_display = {"en": "English", "tr": "Türkçe"}
        self.lang_var = ctk.StringVar(value=lang_display.get(langs.get_language(), "English"))
        ctk.CTkComboBox(lang_col, values=["English", "Türkçe"], variable=self.lang_var, state="readonly", width=120).pack(anchor="w", pady=(2, 0))
        
        theme_col = ctk.CTkFrame(misc_row, fg_color="transparent")
        theme_col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(theme_col, text=t("theme"), anchor="w").pack(fill="x")
        
        theme_options = [t("theme_system"), t("theme_light"), t("theme_dark")]
        saved_theme = self.app_config.get("theme", "System")
        
        # Map internal value to translation string
        display_theme = t("theme_" + saved_theme.lower()) if saved_theme in ["System", "Light", "Dark"] else t("theme_system")
        self.theme_var = ctk.StringVar(value=display_theme)
        
        # When value changes, we can apply instantly, but we'll officially save it in save_settings
        self.theme_cb = ctk.CTkComboBox(theme_col, values=theme_options, variable=self.theme_var, state="readonly", width=120)
        self.theme_cb.pack(anchor="w", pady=(2, 0))
        
        self.delay_var = ctk.StringVar(value=str(self.app_config.get("delay", 2)))
        vcmd_delay = (self.register(self.validate_port), '%P')
        delay_col = ctk.CTkFrame(misc_row, fg_color="transparent")
        delay_col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(delay_col, text=t("delay_label"), anchor="w").pack(fill="x")
        ctk.CTkEntry(delay_col, textvariable=self.delay_var, validate='key', validatecommand=vcmd_delay, width=80).pack(anchor="w", pady=(2, 0))
        
        self.notif_var = ctk.BooleanVar(value=bool(self.app_config.get("notifications", True)))
        notif_col = ctk.CTkFrame(misc_row, fg_color="transparent")
        notif_col.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(notif_col, text=" ", anchor="w").pack(fill="x")
        self.notif_checkbox = ctk.CTkCheckBox(notif_col, text=t("enable_notif"), variable=self.notif_var)
        self.notif_checkbox.pack(anchor="w", pady=(2, 0))
        
        # Row 4: Date + Time Format
        date_map_reverse = {"%d/%m/%Y": "DD/MM/YYYY", "%m/%d/%Y": "MM/DD/YYYY", "%Y-%m-%d": "YYYY-MM-DD", "%d.%m.%Y": "DD.MM.YYYY"}
        time_map_reverse = {"%H:%M": "24H", "%I:%M %p": "12H"}
        
        saved_date = self.app_config.get("date_format", "DD/MM/YYYY")
        if "%" in saved_date: saved_date = date_map_reverse.get(saved_date, "DD/MM/YYYY")
        saved_time = self.app_config.get("time_format", "24H")
        if "%" in saved_time: saved_time = time_map_reverse.get(saved_time, "24H")
        
        datetime_row = ctk.CTkFrame(others_frame, fg_color="transparent")
        datetime_row.pack(fill="x", padx=config.PAD_X, pady=(0, 10))
        
        date_col = ctk.CTkFrame(datetime_row, fg_color="transparent")
        date_col.pack(side="left", fill="x", expand=True)
        self.date_format_var = ctk.StringVar(value=saved_date)
        ctk.CTkLabel(date_col, text=t("date_format"), anchor="w").pack(fill="x")
        date_options = ["DD/MM/YYYY", "MM/DD/YYYY", "YYYY-MM-DD", "DD.MM.YYYY"]
        ctk.CTkComboBox(date_col, values=date_options, variable=self.date_format_var, state="readonly", width=160).pack(anchor="w", pady=(2, 0))
        
        time_col = ctk.CTkFrame(datetime_row, fg_color="transparent")
        time_col.pack(side="left", fill="x", expand=True)
        self.time_format_var = ctk.StringVar(value=saved_time)
        ctk.CTkLabel(time_col, text=t("time_format"), anchor="w").pack(fill="x")
        time_options = ["24H", "12H"]
        ctk.CTkComboBox(time_col, values=time_options, variable=self.time_format_var, state="readonly", width=100).pack(anchor="w", pady=(2, 0))
        
        # Data Folder
        data_folder_frame = ctk.CTkFrame(self.settings_frame)
        data_folder_frame.pack(fill="x", padx=config.PAD_X, pady=(5, 10))
        
        ctk.CTkLabel(data_folder_frame, text=t("data_folder"), font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", padx=config.PAD_X, pady=(10, 0))
        ctk.CTkLabel(data_folder_frame, text=t("data_folder_desc"), anchor="w", text_color="gray", wraplength=600).pack(fill="x", padx=config.PAD_X, pady=(2, 5))
        
        df_row = ctk.CTkFrame(data_folder_frame, fg_color="transparent")
        df_row.pack(fill="x", padx=config.PAD_X, pady=(0, 15))
        
        self.data_folder_var = ctk.StringVar(value=config.get_data_folder_raw())
        ctk.CTkEntry(df_row, textvariable=self.data_folder_var, placeholder_text=config._DEFAULT_APP_DIR, state="readonly").pack(side="left", fill="x", expand=True)
        
        def browse_data_folder():
            folder = filedialog.askdirectory()
            if folder:
                self.data_folder_var.set(folder)
        
        def reset_data_folder():
            self.data_folder_var.set("")
        
        ctk.CTkButton(df_row, text=t("browse"), command=browse_data_folder, width=80).pack(side="left", padx=(5, 0))
        ctk.CTkButton(df_row, text=t("reset_default"), command=reset_data_folder, width=80, fg_color="gray50", hover_color="gray40").pack(side="left", padx=(5, 0))
        
        # --- Signature Frame ---
        sig_frame = ctk.CTkFrame(self.settings_frame)
        sig_frame.pack(fill="x", padx=config.PAD_X, pady=(5, 10))
        
        ctk.CTkLabel(sig_frame, text=t("global_signature") if "global_signature" in langs.TRANSLATIONS.get("en", {}) else "Global Signature", font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", padx=config.PAD_X, pady=(10, 0))
        ctk.CTkLabel(sig_frame, text=t("global_signature_desc") if "global_signature_desc" in langs.TRANSLATIONS.get("en", {}) else "Appended to messages using {signature}. Leave empty to disable.", anchor="w", text_color="gray", wraplength=600).pack(fill="x", padx=config.PAD_X, pady=(2, 5))
        
        sig_inner = ctk.CTkFrame(sig_frame, fg_color="transparent")
        sig_inner.pack(fill="x", padx=config.PAD_X, pady=(0, 15))
        
        self.signature_textbox = ctk.CTkTextbox(sig_inner, height=100)
        self.signature_textbox.pack(fill="x", expand=True)
        self.signature_textbox.insert("1.0", self.app_config.get("signature", ""))
        
        # --- Reports Frame ---
        self.reports_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        report_top_frame = ctk.CTkFrame(self.reports_frame, fg_color="transparent")
        report_top_frame.pack(fill="x", padx=config.PAD_X, pady=20)
        
        self.reports_title = ctk.CTkLabel(report_top_frame, text=t("send_reports"), font=ctk.CTkFont(size=20, weight="bold"))
        self.reports_title.pack(side="left")
        
        self.clear_reports_btn = ctk.CTkButton(report_top_frame, text=t("clear_reports"), fg_color="red", hover_color="darkred", command=self.clear_reports, width=100)
        self.clear_reports_btn.pack(side="right")
        
        self.reports_scrollable_frame = ctk.CTkScrollableFrame(self.reports_frame, label_text=t("recent_logs"))
        self.reports_scrollable_frame.pack(fill="both", expand=True, padx=config.PAD_X, pady=(5, 20))

        # --- Help Frame ---
        self.help_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.help_title = ctk.CTkLabel(self.help_frame, text=t("help_title"), font=ctk.CTkFont(size=24, weight="bold"))
        self.help_title.pack(padx=config.PAD_X, pady=(20, 10), anchor="w")
        
        help_scroll = ctk.CTkScrollableFrame(self.help_frame, fg_color="transparent")
        help_scroll.pack(fill="both", expand=True, padx=config.PAD_X, pady=(0, 20))
        
        def add_help_section(title, body, tb_height=100):
            sec_frame = ctk.CTkFrame(help_scroll)
            sec_frame.pack(fill="x", pady=10)
            ctk.CTkLabel(sec_frame, text=title, font=ctk.CTkFont(size=16, weight="bold"), text_color="#8dbbde").pack(anchor="w", padx=15, pady=(10, 5))
            
            body_tb = ctk.CTkTextbox(sec_frame, height=tb_height, wrap="word", fg_color="transparent")
            body_tb.pack(fill="x", padx=10, pady=(0, 10))
            body_tb.insert("1.0", body)
            body_tb.configure(state="disabled")
            
        add_help_section(t("help_1_title"), t("help_1_body"), 190)
        add_help_section(t("help_2_title"), t("help_2_body"), 160)
        add_help_section(t("help_3_title"), t("help_3_body"), 100)
        add_help_section(t("help_4_title"), t("help_4_body"), 100)

        self.show_contacts_view()

    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
            
    def toggle_all_contacts(self):
        new_state = not self.toggle_all_var.get()
        self.set_all_contacts(new_state)
        
    def set_all_contacts(self, state):
        self.toggle_all_var.set(state)
        for c in self.app_config["contacts"]:
            c["enabled"] = state
            
        self.save_config()
        
        # update UI checkboxes and labels
        for c_id, widgets in self.contact_widgets.items():
            if "is_enabled_var" in widgets:
                widgets["is_enabled_var"].set(state)
                # trigger the toggle command manually for each row frame since setting var doesn't trigger it automatically
                if "toggle_cmd" in widgets:
                    widgets["toggle_cmd"]()
                    
    def toggle_selected_contacts(self):
        if not hasattr(self, 'selected_rows') or not self.selected_rows:
            return
            
        for r in self.selected_rows:
            if r.winfo_exists():
                c_id = getattr(r, "contact_id", None)
                if c_id is not None:
                    widgets = self.contact_widgets.get(c_id)
                    if widgets and "is_enabled_var" in widgets:
                        current_state = widgets["is_enabled_var"].get()
                        widgets["is_enabled_var"].set(not current_state)
                        if "toggle_cmd" in widgets:
                            widgets["toggle_cmd"]()

    def toggle_favorite_selected_contacts(self):
        if not hasattr(self, 'selected_rows') or not self.selected_rows:
            return
            
        for r in self.selected_rows:
            if r.winfo_exists():
                c_id = getattr(r, "contact_id", None)
                if c_id is not None:
                    widgets = self.contact_widgets.get(c_id)
                    if widgets and "toggle_fav_cmd" in widgets:
                        widgets["toggle_fav_cmd"]()

    def delete_selected_contacts(self):
        if not hasattr(self, 'selected_rows') or not self.selected_rows:
            return
            
        if messagebox.askyesno("Delete Selected", f"Are you sure you want to delete {len(self.selected_rows)} selected contacts?"):
            ids_to_del = []
            for r in self.selected_rows:
                if r.winfo_exists() and hasattr(r, "contact_id"):
                    ids_to_del.append(r.contact_id)
                    
            self.app_config["contacts"] = [c for c in self.app_config["contacts"] if c.get("id") not in ids_to_del]
            self.save_config()
            self.refresh_contacts_list()
            self.selected_rows.clear()

    def validate_port(self, P):
        if P == "" or P.isdigit():
            return True
        return False
        
    def validate_email_format(self, event=None):
        email = self.smtp_user_var.get()
        if email and ("@" not in email or "." not in email):
            messagebox.showwarning("Invalid Email", "Please enter a valid email address.")
            
    def toggle_password_visibility(self):
        if self.show_pass_var.get():
            self.pass_entry.configure(show="")
        else:
            self.pass_entry.configure(show="*")

    def load_config(self):
        if os.path.exists(config.CONFIG_FILE):
            try:
                with open(config.CONFIG_FILE, "r") as f:
                    self.app_config = json.load(f)
            except Exception as e:
                print(f"Error loading config: {e}")
                
        self.reports_data = []
        if hasattr(config, "REPORTS_FILE") and os.path.exists(config.REPORTS_FILE):
            try:
                with open(config.REPORTS_FILE, "r") as f:
                    self.reports_data = json.load(f)
            except Exception as e:
                print(f"Error loading reports: {e}")

        if "smtp" not in self.app_config:
            self.app_config["smtp"] = {}
        if "contacts" not in self.app_config:
            self.app_config["contacts"] = []
        if "delay" not in self.app_config:
            self.app_config["delay"] = 2
        if "notifications" not in self.app_config:
            self.app_config["notifications"] = True
        if "signature" not in self.app_config:
            self.app_config["signature"] = ""

        for i, c in enumerate(self.app_config["contacts"]):
            if "id" not in c:
                c["id"] = i
            if "enabled" not in c:
                c["enabled"] = True
            if "_status_lbl_ref" in c:
                del c["_status_lbl_ref"]

    def save_config(self):
        with open(config.CONFIG_FILE, "w") as f:
            json.dump(self.app_config, f, indent=4)
            
    def save_reports(self):
        if hasattr(config, "REPORTS_FILE"):
            with open(config.REPORTS_FILE, "w") as f:
                json.dump(self.reports_data, f, indent=4)

    def show_contacts_view(self):
        self.settings_frame.grid_forget()
        self.reports_frame.grid_forget()
        self.help_frame.grid_forget()
        self.favorites_frame.grid_forget()
        self.contacts_frame.grid(row=0, column=1, sticky="nsew")
        self.send_all_button.grid(row=7, column=0, padx=15, pady=(15, 25))
        
        # Only refresh if the widgets haven't been created yet to prevent lag
        if not self.contact_widgets and self.app_config.get("contacts"):
            self.refresh_contacts_list()

    def show_settings_view(self):
        self.contacts_frame.grid_forget()
        self.reports_frame.grid_forget()
        self.help_frame.grid_forget()
        self.favorites_frame.grid_forget()
        self.send_all_button.grid_remove()
        self.settings_frame.grid(row=0, column=1, sticky="nsew")
        
    def show_reports_view(self):
        self.contacts_frame.grid_forget()
        self.settings_frame.grid_forget()
        self.help_frame.grid_forget()
        self.favorites_frame.grid_forget()
        self.send_all_button.grid_remove()
        self.reports_frame.grid(row=0, column=1, sticky="nsew")
        self.refresh_reports_list()

    def show_help_view(self):
        self.contacts_frame.grid_forget()
        self.settings_frame.grid_forget()
        self.reports_frame.grid_forget()
        self.favorites_frame.grid_forget()
        self.send_all_button.grid_remove()
        self.help_frame.grid(row=0, column=1, sticky="nsew")

    def show_favorites_view(self):
        self.contacts_frame.grid_forget()
        self.settings_frame.grid_forget()
        self.reports_frame.grid_forget()
        self.help_frame.grid_forget()
        self.send_all_button.grid_remove()
        self.favorites_frame.grid(row=0, column=1, sticky="nsew")
        self.refresh_favorites_list()

    def refresh_favorites_list(self):
        for w in self.fav_list_frame.winfo_children():
            w.destroy()
            
        favs = [c for c in self.app_config.get("contacts", []) if c.get("favorite")]
        
        if not favs:
            ctk.CTkLabel(self.fav_list_frame, text=t("no_favorite_selected"), text_color="gray").pack(pady=20)
            self.fav_title_lbl.configure(text=t("no_favorite_selected"))
            self.fav_notes_textbox.configure(state="normal")
            self.fav_notes_textbox.delete("1.0", "end")
            self.fav_notes_textbox.configure(state="disabled")
            self.active_fav_contact_id = None
            return
            
        def select_fav(contact_id):
            self.active_fav_contact_id = contact_id
            c = next((x for x in self.app_config["contacts"] if x.get("id") == contact_id), None)
            if c:
                name = c.get("company") or c.get("email")
                self.fav_title_lbl.configure(text=f"{t('notes')} - {name}")
                
                self.fav_notes_textbox.configure(state="normal")
                self.fav_notes_textbox.delete("1.0", "end")
                self.fav_notes_textbox.insert("1.0", c.get("notes", ""))
                
                # Highlight active row in the sidebar
                for child in self.fav_list_frame.winfo_children():
                    child.configure(fg_color="#1f538d" if getattr(child, "c_id", None) == contact_id else "transparent")
                    
        for c in favs:
            row_frame = ctk.CTkFrame(self.fav_list_frame, fg_color="transparent", cursor="hand2")
            row_frame.pack(fill="x", pady=2, padx=5)
            row_frame.c_id = c.get("id")
            
            name = c.get("company") or c.get("email")
            lbl = ctk.CTkLabel(row_frame, text=name, anchor="w", font=ctk.CTkFont(weight="bold"))
            lbl.pack(fill="x", padx=10, pady=8)
            
            for w in [row_frame, lbl]:
                w.bind("<Button-1>", lambda e, cid=c.get("id"): select_fav(cid))

    def clear_reports(self):
        if messagebox.askyesno(t("clear_reports"), t("clear_reports_confirm")):
            self.reports_data = []
            self.save_reports()
            self.refresh_reports_list()
            
    def toggle_report_group(self, subframe):
        if subframe.winfo_ismapped():
            subframe.pack_forget()
        else:
            subframe.pack(fill="x", padx=10, pady=(0, 5))
            
    def refresh_reports_list(self):
        for widget in self.reports_scrollable_frame.winfo_children():
            widget.destroy()
            
        for run in reversed(self.reports_data):
            run_frame = ctk.CTkFrame(self.reports_scrollable_frame, fg_color="transparent")
            run_frame.pack(fill="x", pady=5, padx=5)
            
            subframe = ctk.CTkFrame(run_frame, fg_color=("gray85", "gray20"))
            
            deliveries = [r for r in run.get("deliveries", []) if r.get("status") != "Skipped (Disabled)"]
            if not deliveries:
                continue

            summary_text = t("send_batch", date=run.get('date', 'Unknown Time'), count=len(deliveries))
            
            header_frame = ctk.CTkFrame(run_frame, fg_color="transparent")
            header_frame.pack(fill="x")
            
            btn = ctk.CTkButton(header_frame, text=summary_text, anchor="w", fg_color="#1f538d",
                                command=lambda sf=subframe: self.toggle_report_group(sf))
            btn.pack(side="left", fill="x", expand=True)
            
            def delete_batch(r=run):
                if messagebox.askyesno(t("delete_confirm_title"), t("delete_batch_confirm", date=r.get('date'))):
                    self.reports_data.remove(r)
                    self.save_reports()
                    self.refresh_reports_list()
                    
            del_btn = ctk.CTkButton(header_frame, text="X", width=30, fg_color="transparent", text_color="#ff5555", hover_color="#8b0000", command=delete_batch)
            del_btn.pack(side="right", padx=(5, 0))
            
            for rep in deliveries:
                row = ctk.CTkFrame(subframe, fg_color="transparent", cursor="hand2")
                row.pack(fill="x", pady=2, padx=5)
                
                row.grid_columnconfigure(0, weight=1)
                row.grid_columnconfigure(1, weight=1)
                row.grid_columnconfigure(2, weight=1)
                
                email_lbl = ctk.CTkLabel(row, text=rep.get("email", ""), anchor="w")
                email_lbl.grid(row=0, column=0, sticky="ew", padx=10)
                
                subj_lbl = ctk.CTkLabel(row, text=rep.get("subject", ""), anchor="w")
                subj_lbl.grid(row=0, column=1, sticky="ew", padx=10)
                
                status_txt = rep.get("status", "")
                color = "green" if status_txt == "Sent" else ("gray" if "Skip" in status_txt else "red")
                status_lbl = ctk.CTkLabel(row, text=status_txt, text_color=color, anchor="e")
                status_lbl.grid(row=0, column=2, sticky="ew", padx=10)

                def show_msg(event, r=rep):
                    self.show_report_message_popup(r)
                
                for w in [row, email_lbl, subj_lbl, status_lbl]:
                    w.bind("<Button-1>", show_msg)

    def show_report_message_popup(self, rep_data):
        if getattr(self, 'report_popup', None) and self.report_popup.winfo_exists():
            self.report_popup.destroy()
            
        self.report_popup = ctk.CTkToplevel(self)
        self.report_popup.title(t("message_details"))
        self.report_popup.geometry("500x450")
        self.report_popup.attributes("-topmost", True)
        self.report_popup.transient(self)
        
        self.report_popup.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (500 // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (450 // 2)
        self.report_popup.geometry(f"+{x}+{y}")
        
        info_frame = ctk.CTkFrame(self.report_popup, fg_color="transparent")
        info_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(info_frame, text=f"{t('report_to')}: {rep_data.get('email', '')}", font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x")
        ctk.CTkLabel(info_frame, text=f"{t('report_subject')}: {rep_data.get('subject', '')}", font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x", pady=(5,0))
        
        msg_frame = ctk.CTkFrame(self.report_popup)
        msg_frame.pack(fill="both", expand=True, padx=20, pady=(5, 10))
        
        textbox = ctk.CTkTextbox(msg_frame)
        textbox.pack(fill="both", expand=True, padx=2, pady=2)
        textbox.insert("1.0", rep_data.get("message", t("no_message")))
        textbox.configure(state="disabled")
        
        att_str = rep_data.get('attachment', '')
        if att_str:
            att_frame = ctk.CTkFrame(self.report_popup, fg_color="transparent")
            att_frame.pack(fill="x", padx=20, pady=(0, 20))
            ctk.CTkLabel(att_frame, text=t("attachments_label"), font=ctk.CTkFont(weight="bold"), anchor="w", text_color="lightgray").pack(fill="x")
            
            att_list = [a.strip() for a in att_str.split(",")]
            for a in att_list:
                if a:
                    ctk.CTkLabel(att_frame, text=f"- {a}", anchor="w", text_color="lightgray").pack(fill="x", padx=10)

    def save_settings(self):
        try:
            delay = int(self.delay_var.get().strip())
        except ValueError:
            delay = 2
            
        self.app_config["smtp"] = {
            "server": self.smtp_server_var.get().strip(),
            "port": self.smtp_port_var.get().strip(),
            "user": self.smtp_user_var.get().strip(),
            "password": crypto.encrypt(self.smtp_pass_var.get().strip())
        }
        self.app_config["delay"] = delay
        self.app_config["date_format"] = self.date_format_var.get().strip()
        self.app_config["time_format"] = self.time_format_var.get().strip()
        self.app_config["notifications"] = self.notif_var.get()
        self.app_config["signature"] = self.signature_textbox.get("1.0", "end-1c").strip()
        
        # Save language
        lang_reverse = {"English": "en", "Türkçe": "tr"}
        new_lang = lang_reverse.get(self.lang_var.get(), "en")
        old_lang = self.app_config.get("language", "")
        self.app_config["language"] = new_lang
        langs.set_language(new_lang)

        # Save Theme
        theme_display = self.theme_var.get()
        # Reverse translation mapping
        if theme_display == t("theme_light"): new_theme = "Light"
        elif theme_display == t("theme_dark"): new_theme = "Dark"
        else: new_theme = "System"
        
        self.app_config["theme"] = new_theme
        ctk.set_appearance_mode(new_theme)
        
        # Handle data folder change
        new_data_folder = self.data_folder_var.get().strip()
        old_data_folder = config.get_data_folder_raw()
        need_restart = False
        
        if new_data_folder != old_data_folder:
            old_app_dir = config.APP_DIR  # Current data directory before change
            
            # Set the new data folder pointer and reload paths
            config.set_data_folder(new_data_folder)
            config.reload_paths()
            new_app_dir = config.APP_DIR  # New data directory
            
            # Migrate data files from old to new location
            files_to_move = ["config.json", "reports.json"]
            dirs_to_move = ["attachments"]
            
            for fname in files_to_move:
                src = os.path.join(old_app_dir, fname)
                dst = os.path.join(new_app_dir, fname)
                if os.path.exists(src) and not os.path.exists(dst):
                    try:
                        shutil.move(src, dst)
                    except Exception as e:
                        print(f"Failed to move {fname}: {e}")
            
            for dname in dirs_to_move:
                src = os.path.join(old_app_dir, dname)
                dst = os.path.join(new_app_dir, dname)
                if os.path.isdir(src) and not os.path.exists(dst):
                    try:
                        shutil.move(src, dst)
                    except Exception as e:
                        print(f"Failed to move {dname}: {e}")
            
            messagebox.showinfo(t("info"), t("data_folder_moved"))
            need_restart = True
        
        self.save_config()
        
        if not need_restart and new_lang != old_lang:
            need_restart = True
        
        if not need_restart:
            messagebox.showinfo(t("success"), t("settings_saved"))
        
        if need_restart:
            self.withdraw()
            # Strip existing --view args to prevent accumulation
            clean_args = []
            skip_next = False
            for arg in sys.argv:
                if skip_next:
                    skip_next = False
                    continue
                if arg == "--view":
                    skip_next = True
                    continue
                clean_args.append(arg)
            os.execv(sys.executable, [sys.executable] + clean_args + ["--view", "settings"])

    def apply_sorting(self, event=None):
        # We no longer destroy all widgets; instead we just un-pack and re-pack them in order.
        self.filter_contacts_list()

    def filter_contacts_list(self, event=None):
        query = self.search_var.get().strip().lower()
        sort_mode = self.sort_var.get()
        
        # Step 1: Hide all current widgets so we can re-pack them in the new order
        for widgets in self.contact_widgets.values():
            if "row_frame" in widgets:
                widgets["row_frame"].pack_forget()

        # Step 2: Determine order
        contacts_to_render = self.app_config["contacts"]
        if getattr(self, "sort_name_var", None) and self.sort_name_var.get():
            contacts_to_render = sorted(contacts_to_render, key=lambda x: x.get("company", "").lower())
        
        # Step 3: Filter and pack visible ones
        for c in contacts_to_render:
            widgets = self.contact_widgets.get(c.get("id"))
            if not widgets or "row_frame" not in widgets:
                continue
                
            row_frame = widgets["row_frame"]
            
            show_item = True
            if sort_mode == "active" and not c.get("enabled", True):
                show_item = False
            elif sort_mode == "inactive" and c.get("enabled", True):
                show_item = False
            elif sort_mode == "favs" and not c.get("favorite", False):
                show_item = False
                
            tag_mode = getattr(self, "tag_filter_var", None)
            if show_item and tag_mode and tag_mode.get() != t("all_tags"):
                if c.get("tag", "").strip() != tag_mode.get():
                    show_item = False
                
            if show_item and query:
                comp = c.get("company", "").lower()
                email = c.get("email", "").lower()
                if query not in comp and query not in email:
                    show_item = False
            
            if show_item:
                row_frame.pack(fill="x", pady=5, padx=5)

    def refresh_contacts_list(self):
        # This function should only be called when loading config, adding, or deleting contacts.
        for widget in self.contacts_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.contact_widgets.clear()
        
        # Update available tags dropdown based on current data
        if hasattr(self, "btn_filter_tag"):
            existing_tags = sorted(list(set(c.get("tag", "").strip() for c in self.app_config["contacts"] if c.get("tag", "").strip())))
            self.btn_filter_tag.configure(values=[t("all_tags")] + existing_tags)
            if self.tag_filter_var.get() not in [t("all_tags")] + existing_tags:
                self.tag_filter_var.set(t("all_tags"))
        
        for c in self.app_config["contacts"]:
            self.create_contact_row(c)
            
        # Apply filters/sort immediately after drawing
        self.filter_contacts_list()

    def import_csv(self):
        filepath = filedialog.askopenfilename(title="Select CSV to Import", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not filepath:
            return
            
        try:
            with open(filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                headers = reader.fieldnames or []
                
                # Try to map common header names, fallback to empty string if not found
                def get_val(row, possible_keys):
                    for k in possible_keys:
                        for actual_k in headers:
                            if actual_k and k.lower() == actual_k.lower().strip():
                                return row[actual_k].strip()
                    return ""

                added_count = 0
                for row in reader:
                    email_val = get_val(row, ['email', 'e-mail', 'mail'])
                    if not email_val:
                        continue # Skip rows without email
                        
                    new_id = 0 if not self.app_config["contacts"] else max(int(c.get("id", 0)) for c in self.app_config["contacts"]) + 1
                    
                    new_contact = {
                        "id": new_id,
                        "company": get_val(row, ['company', 'name', 'client', 'contact']),
                        "email": email_val,
                        "subject": get_val(row, ['subject', 'title']),
                        "message": get_val(row, ['message', 'body', 'text', 'content']),
                        "attachments": [],
                        "enabled": True
                    }
                    self.app_config["contacts"].append(new_contact)
                    added_count += 1

            if added_count > 0:
                self.save_config()
                self.refresh_contacts_list()
                messagebox.showinfo(t("success"), t("csv_import_success", count=added_count))
            else:
                messagebox.showwarning(t("info"), t("csv_no_email"))
                
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import CSV:\n{str(e)}")

    def export_csv(self):
        if not self.app_config["contacts"]:
            messagebox.showinfo(t("info"), t("csv_export_success", count=0))
            return
            
        filepath = filedialog.asksaveasfilename(title="Export to CSV", defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not filepath:
            return
            
        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                fieldnames = ['company', 'email', 'subject', 'message']
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for c in self.app_config["contacts"]:
                    writer.writerow({
                        'company': c.get('company', ''),
                        'email': c.get('email', ''),
                        'subject': c.get('subject', ''),
                        'message': c.get('message', '')
                    })
            messagebox.showinfo(t("success"), t("csv_export_success", count=len(self.app_config['contacts'])))
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export CSV:\n{str(e)}")

    def create_contact_row(self, contact):
        c_id = contact.get("id")
        c_enabled = contact.get("enabled", True)
        c_favorite = contact.get("favorite", False)
        c_company = contact.get("company", "")
        c_email = contact.get("email", "")
        c_subject = contact.get("subject", "")
        c_tag = contact.get("tag", "").strip()
        
        row_frame = ctk.CTkFrame(self.contacts_scrollable_frame)
        row_frame.pack(fill="x", pady=5, padx=5)
        
        row_frame.contact_id = c_id
        
        row_frame.grid_columnconfigure(1, weight=0)
        row_frame.grid_columnconfigure(2, weight=0) # fav_btn
        row_frame.grid_columnconfigure(3, weight=1, uniform="col") # company
        row_frame.grid_columnconfigure(4, weight=2, uniform="col") # email
        row_frame.grid_columnconfigure(5, weight=2, uniform="col") # subject
        row_frame.grid_columnconfigure(6, weight=0) # status
        row_frame.grid_columnconfigure(7, weight=0) # edit
        row_frame.grid_columnconfigure(8, weight=0) # del
        
        drag_handle = ctk.CTkLabel(row_frame, text="", cursor="fleur", width=0)
        drag_handle.grid(row=0, column=0, padx=(5, 0), pady=10)

        is_enabled = ctk.BooleanVar(value=c_enabled)
        
        enable_cb = ctk.CTkCheckBox(row_frame, text="", variable=is_enabled, width=20)
        enable_cb.grid(row=0, column=1, padx=5)
        
        is_fav = ctk.BooleanVar(value=c_favorite)
        fav_btn = ctk.CTkButton(row_frame, text="⭐️" if c_favorite else "☆", width=30, fg_color="transparent", text_color="gold" if c_favorite else "gray", hover_color="gray30", font=ctk.CTkFont(size=18))
        fav_btn.grid(row=0, column=2, padx=5)
        
        def toggle_favorite():
            new_state = not is_fav.get()
            is_fav.set(new_state)
            fav_btn.configure(text="⭐️" if new_state else "☆", text_color="gold" if new_state else "gray")
            
            for c in self.app_config["contacts"]:
                if c.get("id") == c_id:
                    c["favorite"] = new_state
                    self.save_config()
                    break
                    
        fav_btn.configure(command=toggle_favorite)

        def truncate(text, max_len=20):
            return text if len(text) <= max_len else text[:max_len-3] + "..."
            
        font_style = ctk.CTkFont(overstrike=not c_enabled, slant="italic" if not c_enabled else "roman")
        text_color = "gray50" if not c_enabled else ["gray10", "#DCE4EE"]

        display_name = truncate(c_company)
        if c_tag:
            display_name = f"[{c_tag}] {display_name}"
            
        company_lbl = ctk.CTkLabel(row_frame, text=display_name, anchor="w", cursor="fleur", text_color=text_color, font=font_style)
        company_lbl.grid(row=0, column=3, sticky="w", padx=5)
        
        email_lbl = ctk.CTkLabel(row_frame, text=truncate(c_email, 25), anchor="w", cursor="fleur", text_color=text_color, font=font_style)
        email_lbl.grid(row=0, column=4, sticky="w", padx=5)
        
        subj_lbl = ctk.CTkLabel(row_frame, text=truncate(c_subject, 25), anchor="w", cursor="fleur", text_color=text_color, font=font_style)
        subj_lbl.grid(row=0, column=5, sticky="w", padx=5)

        status_lbl = ctk.CTkLabel(row_frame, text=t("ready"), width=80)
        status_lbl.grid(row=0, column=6, padx=10)
        
        self.contact_widgets[c_id] = {
            "status_lbl": status_lbl,
            "is_enabled_var": is_enabled,
            "row_frame": row_frame,
            "toggle_fav_cmd": toggle_favorite
        }
        
        edit_btn = ctk.CTkButton(row_frame, text=t("edit"), width=60, command=lambda c=contact: self.open_contact_popup(c))
        edit_btn.grid(row=0, column=7, padx=5)

        del_btn = ctk.CTkButton(row_frame, text=t("del"), width=60, fg_color="red", hover_color="darkred", command=lambda c=contact: self.delete_contact(c))
        del_btn.grid(row=0, column=8, padx=(5, 10))
        
        def toggle_contact_state():
            enabled = is_enabled.get()
            new_f_style = ctk.CTkFont(overstrike=not enabled, slant="italic" if not enabled else "roman")
            new_t_color = "gray50" if not enabled else ["gray10", "#DCE4EE"]
            company_lbl.configure(text_color=new_t_color, font=new_f_style)
            email_lbl.configure(text_color=new_t_color, font=new_f_style)
            subj_lbl.configure(text_color=new_t_color, font=new_f_style)
            
            for c in self.app_config["contacts"]:
                if c.get("id") == c_id:
                    c["enabled"] = enabled
                    self.save_config()
                    break

        enable_cb.configure(command=toggle_contact_state)
        self.contact_widgets[c_id]["toggle_cmd"] = toggle_contact_state
        
        for w in [row_frame, drag_handle, company_lbl, email_lbl, subj_lbl]:
            w.bind("<Button-1>", lambda e, r=row_frame: self.start_drag(e, r))
            w.bind("<B1-Motion>", self.do_drag)
            w.bind("<ButtonRelease-1>", self.stop_drag)
            w.bind("<Button-3>", self.show_context_menu)
            
        def select_row(event, rf=row_frame):
            self.select_contact_row(event, rf)
            
        for w in [row_frame, company_lbl, email_lbl, subj_lbl]:
            w.bind("<Button-1>", select_row, add="+")
            
    def select_contact_row(self, event, row_frame):
        self.pending_single_select = None
        
        if not hasattr(self, 'selected_rows'):
            self.selected_rows = []
            
        ctrl_pressed = (event.state & 0x0004) != 0
        shift_pressed = (event.state & 0x0001) != 0

        if ctrl_pressed:
            if row_frame in self.selected_rows:
                self.selected_rows.remove(row_frame)
                row_frame.configure(border_width=0)
            else:
                self.selected_rows.append(row_frame)
                row_frame.configure(border_width=2, border_color=config.ACCENT_COLOR)
            self.last_selected_row = row_frame
            
        elif shift_pressed and getattr(self, 'last_selected_row', None) and self.last_selected_row.winfo_exists():
            children = [w for w in self.contacts_scrollable_frame.winfo_children() if hasattr(w, "contact_id")]
            # Sort by visual position so the slice matches what the user sees
            children.sort(key=lambda w: w.winfo_y())
            try:
                start_idx = children.index(self.last_selected_row)
                end_idx = children.index(row_frame)
                
                if start_idx > end_idx:
                    start_idx, end_idx = end_idx, start_idx
                    
                for r in self.selected_rows:
                    if r.winfo_exists():
                        r.configure(border_width=0)
                self.selected_rows.clear()
                
                for i in range(start_idx, end_idx + 1):
                    r = children[i]
                    self.selected_rows.append(r)
                    r.configure(border_width=2, border_color=config.ACCENT_COLOR)
                    
            except ValueError:
                pass
        else:
            if row_frame in self.selected_rows and len(self.selected_rows) > 1:
                # Instead of clearing selection instantly, defer it in case they want to drag
                self.pending_single_select = row_frame
            else:
                for r in self.selected_rows:
                    if r.winfo_exists():
                        r.configure(border_width=0)
                self.selected_rows = [row_frame]
                row_frame.configure(border_width=2, border_color=config.ACCENT_COLOR)
                self.last_selected_row = row_frame
        
        self.bind("<Delete>", self.delete_selected_contact)
        
    def delete_selected_contact(self, event=None):
        if hasattr(self, 'selected_rows') and self.selected_rows:
            to_delete = []
            for r in self.selected_rows:
                if r.winfo_exists():
                    c_id = getattr(r, "contact_id", None)
                    if c_id is not None:
                        to_delete.append(c_id)
            
            if to_delete:
                msg = t("delete_selected_confirm", count=len(to_delete)) if len(to_delete) > 1 else t("delete_confirm_msg", email="")
                if messagebox.askyesno(t("delete_confirm_title"), msg):
                    self.app_config["contacts"] = [c for c in self.app_config["contacts"] if c.get("id") not in to_delete]
                    self.save_config()
                    self.selected_rows.clear()
                    self.refresh_contacts_list()

    def start_drag(self, event, row_frame):
        self.drag_start_y = event.y_root
        self.potential_drag_row = row_frame
        self.is_dragging = False

    def init_drag_visuals(self):
        row_frame = self.potential_drag_row
        
        if not hasattr(self, 'selected_rows'):
            self.selected_rows = []
            
        # If the user drags a row that isn't selected, they intend to drag just that row
        if row_frame not in self.selected_rows:
            for r in self.selected_rows:
                if r.winfo_exists(): r.configure(border_width=0)
            self.selected_rows = [row_frame]
            row_frame.configure(border_width=2, border_color=config.ACCENT_COLOR)
            self.last_selected_row = row_frame

        self.dragged_rows = [r for r in self.selected_rows if r.winfo_exists()]
        self.dragged_rows.sort(key=lambda r: r.winfo_y())

        self.drag_win = ctk.CTkToplevel(self)
        self.drag_win.overrideredirect(True)
        self.drag_win.attributes("-alpha", 0.8)
        self.drag_win.attributes("-topmost", True)
        
        info_text = t("moving_contacts", n=len(self.dragged_rows))
        
        drag_frame = ctk.CTkFrame(self.drag_win, fg_color="#1f538d", corner_radius=5)
        drag_frame.pack(fill="both", expand=True)
        ctk.CTkLabel(drag_frame, text=info_text, padx=10, pady=5, font=ctk.CTkFont(weight="bold")).pack()
        
        self.drag_offset_x = 15
        self.drag_offset_y = 15
        self.drag_win.geometry(f"+{self.winfo_pointerx() + self.drag_offset_x}+{self.winfo_pointery() + self.drag_offset_y}")
        
        total_height = sum(r.winfo_height() + 10 for r in self.dragged_rows) # +10 for padding
        self.placeholder = ctk.CTkFrame(self.contacts_scrollable_frame, height=total_height, fg_color="transparent", border_width=2, border_color="#1f538d")
        
        first_row = self.dragged_rows[0]
        self.placeholder.pack(before=first_row, fill="x", pady=5, padx=5)
        
        for r in self.dragged_rows:
            r.pack_forget()

    def do_drag(self, event):
        if not getattr(self, 'is_dragging', False):
            if hasattr(self, 'drag_start_y') and abs(event.y_root - self.drag_start_y) > 5:
                self.is_dragging = True
                self.pending_single_select = None
                self.init_drag_visuals()
                self.last_drag_y = event.y_root
            return

        if not getattr(self, 'dragged_rows', None) or not getattr(self, 'placeholder', None):
            return
            
        if hasattr(self, 'drag_win') and self.drag_win.winfo_exists():
            self.drag_win.geometry(f"+{event.x_root + self.drag_offset_x}+{event.y_root + self.drag_offset_y}")
        
        if abs(event.y_root - getattr(self, 'last_drag_y', event.y_root)) < 5:
            return
        self.last_drag_y = event.y_root
            
        y = event.y_root
        
        children = [w for w in self.contacts_scrollable_frame.winfo_children() if hasattr(w, "contact_id")]
        
        for widget in children:
            if widget in self.dragged_rows:
                continue
            
            wy = widget.winfo_rooty()
            wh = widget.winfo_height()
            
            if wy < y < wy + wh:
                if y < wy + (wh / 2):
                    self.placeholder.pack(before=widget)
                else:
                    self.placeholder.pack(after=widget)
                break

    def stop_drag(self, event):
        if not getattr(self, 'is_dragging', False):
            if getattr(self, 'pending_single_select', None):
                row_frame = self.pending_single_select
                for r in self.selected_rows:
                    if r.winfo_exists():
                        r.configure(border_width=0)
                self.selected_rows = [row_frame]
                if row_frame.winfo_exists():
                    row_frame.configure(border_width=2, border_color=config.ACCENT_COLOR)
                self.last_selected_row = row_frame
                self.pending_single_select = None
            return
            
        self.pending_single_select = None
            
        if hasattr(self, 'drag_win') and getattr(self, 'drag_win', None) and self.drag_win.winfo_exists():
            self.drag_win.destroy()
            self.drag_win = None
            
        if hasattr(self, 'placeholder') and getattr(self, 'placeholder', None) and self.placeholder.winfo_exists():
            # Repack all dragged rows after placeholder sequentially
            for r in reversed(self.dragged_rows):
                r.pack(after=self.placeholder, fill="x", pady=5, padx=5)
            self.placeholder.destroy()
            self.placeholder = None
        elif getattr(self, 'dragged_rows', None):
            for r in self.dragged_rows:
                r.pack(fill="x", pady=5, padx=5)
            
        self.dragged_rows = None
        self.is_dragging = False
        
        # Ensure the layout is fully redrawn before reading winfo_y()
        self.update_idletasks()
        
        new_contacts = []
        # winfo_children() returns in creation order, not pack order. We must sort by Y coordinate to get visual order.
        children_widgets = [w for w in self.contacts_scrollable_frame.winfo_children() if hasattr(w, "contact_id")]
        children_widgets.sort(key=lambda w: w.winfo_y())
        
        for widget in children_widgets:
            c_id = getattr(widget, "contact_id", None)
            if c_id is not None:
                contact = next((c for c in self.app_config["contacts"] if c.get("id") == c_id), None)
                if contact:
                    new_contacts.append(contact)
                    
        if len(new_contacts) == len(self.app_config["contacts"]):
            self.app_config["contacts"] = new_contacts
            self.save_config()

    def delete_contact(self, contact_to_del):
        if messagebox.askyesno(t("delete_confirm_title"), t("delete_confirm_msg", email=contact_to_del.get('email'))):
            self.app_config["contacts"] = [c for c in self.app_config["contacts"] if c.get("id") != contact_to_del.get("id")]
            self.save_config()
            self.refresh_contacts_list()

    def open_contact_popup(self, contact=None):
        popup = ctk.CTkToplevel(self)
        popup.title(t("add_new_contact_title") if contact is None else t("edit_contact_title"))
        popup.geometry("920x560")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        icon_path = resource_path(os.path.join("assets", "logo.ico"))
        if os.path.exists(icon_path):
            popup.iconbitmap(icon_path)

        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 920) // 2
        y = self.winfo_y() + (self.winfo_height() - 560) // 2
        popup.geometry(f"+{x}+{y}")

        comp_var = ctk.StringVar(value=contact.get("company", "") if contact else "")
        email_var = ctk.StringVar(value=contact.get("email", "") if contact else "")
        subj_var = ctk.StringVar(value=contact.get("subject", "") if contact else "")
        tag_var = ctk.StringVar(value=contact.get("tag", "") if contact else "")
        
        legacy_att = contact.get("attachment", "") if contact else ""
        existing_atts = contact.get("attachments", [legacy_att] if legacy_att else []) if contact else []
        existing_atts = [resolve_att_path(p) for p in existing_atts]
        files_var = ctk.StringVar(value="|".join(existing_atts))

        def save_changes():
            email_val = email_var.get().strip()
            if not email_val:
                messagebox.showerror(t("error"), t("email_required"), parent=popup)
                return

            final_attachments = [p for p in files_var.get().split("|") if p]

            if contact:
                c_id = contact.get("id")
                for c in self.app_config["contacts"]:
                    if c.get("id") == c_id:
                        c["company"] = comp_var.get().strip()
                        c["email"] = email_val
                        c["subject"] = subj_var.get().strip()
                        c["tag"] = tag_var.get().strip()
                        c["message"] = msg_textbox.get("1.0", "end-1c").strip()
                        c["attachments"] = final_attachments
                        if "attachment" in c: del c["attachment"]
                        break
            else:
                new_id = 0 if not self.app_config["contacts"] else max(int(c.get("id", 0)) for c in self.app_config["contacts"]) + 1
                new_contact = {
                    "id": new_id,
                    "company": comp_var.get().strip(),
                    "email": email_val,
                    "subject": subj_var.get().strip(),
                    "tag": tag_var.get().strip(),
                    "message": msg_textbox.get("1.0", "end-1c").strip(),
                    "attachments": final_attachments,
                    "enabled": True
                }
                self.app_config["contacts"].insert(0, new_contact)

            self.save_config()
            self.refresh_contacts_list()
            popup.destroy()

        bottom_bar = ctk.CTkFrame(popup, fg_color="transparent")
        bottom_bar.pack(side="bottom", fill="x", padx=20, pady=(5, 15))
        
        save_btn_bottom = ctk.CTkButton(bottom_bar, text=t("save_contact_list") if contact else t("add_to_contact_list"), command=save_changes, fg_color="#1f538d", height=40, font=ctk.CTkFont(size=14, weight="bold"))
        save_btn_bottom.pack(side="right", padx=5)

        main_frame = ctk.CTkFrame(popup, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=(15, 5))
        
        left_frame = ctk.CTkFrame(main_frame, fg_color="transparent", width=300)
        left_frame.pack(side="left", fill="both", expand=False, padx=(0, 10))
        left_frame.pack_propagate(False)
        
        right_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        right_frame.pack(side="right", fill="both", expand=True)

        ctk.CTkLabel(left_frame, text=t("company_name"), anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(left_frame, textvariable=comp_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        ctk.CTkLabel(left_frame, text=t("recipient_email"), anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(left_frame, textvariable=email_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        ctk.CTkLabel(left_frame, text=t("subject"), anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(left_frame, textvariable=subj_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        ctk.CTkLabel(left_frame, text=t("tag"), anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        existing_tags = sorted(list(set(c.get("tag", "").strip() for c in self.app_config["contacts"] if c.get("tag", "").strip())))
        if not existing_tags: existing_tags = [""]
        ctk.CTkComboBox(left_frame, variable=tag_var, values=existing_tags).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        ctk.CTkLabel(right_frame, text=t("message_body"), anchor="w", font=ctk.CTkFont(weight="bold")).pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        
        vars_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        vars_frame.pack(fill="x", padx=5)
        ctk.CTkLabel(vars_frame, text=t("insert_label"), font=ctk.CTkFont(size=11, slant="italic"), text_color="gray").pack(side="left", padx=(0, 5))
        
        template_vars = ["{company_name}", "{email}", "{email_prefix}", "{date}", "{time}", "{signature}"]
        for var in template_vars:
            def make_insert(v=var):
                def _insert():
                    msg_textbox.insert(tk.INSERT, v)
                return _insert
            ctk.CTkButton(vars_frame, text=var, width=len(var)*8, height=22, fg_color="gray30", hover_color="gray40", font=ctk.CTkFont(size=11), command=make_insert()).pack(side="left", padx=2)
        
        msg_textbox = ctk.CTkTextbox(right_frame)
        msg_textbox.pack(fill="both", expand=True, padx=5, pady=config.ENTRY_PAD_Y)
        if contact:
            msg_textbox.insert("1.0", contact.get("message", ""))
            
        # Message right-click paste/copy menu
        msg_menu = tk.Menu(popup, tearoff=0)
        
        def copy_text():
            try:
                selected_text = ""
                try: selected_text = msg_textbox.selection_get()
                except tk.TclError: pass
                
                popup.clipboard_clear()
                popup.clipboard_append(selected_text)
            except Exception:
                pass
                
        def paste_text():
            try:
                msg_textbox.insert(tk.INSERT, popup.clipboard_get())
            except Exception:
                pass
                
        msg_menu.add_command(label=t("copy"), command=copy_text)
        msg_menu.add_command(label=t("paste"), command=paste_text)
        
        def show_msg_menu(event):
            try:
                msg_menu.tk_popup(event.x_root, event.y_root)
            finally:
                msg_menu.grab_release()

        msg_textbox._textbox.bind("<Button-3>", show_msg_menu)

        ctk.CTkLabel(left_frame, text=t("attachments"), anchor="w", font=ctk.CTkFont(weight="bold")).pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        
        prog_lbl = ctk.CTkLabel(left_frame, text="", text_color="gray")
        prog_bar = ctk.CTkProgressBar(left_frame)
        prog_bar.set(0)
        
        file_list_frame = ctk.CTkScrollableFrame(left_frame, fg_color=("gray85", "gray20"))
        file_list_frame.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        def update_files_display(paths_list):
            for w in file_list_frame.winfo_children():
                w.destroy()
            
            valid_paths = [p for p in paths_list if p]
            if not valid_paths:
                ctk.CTkLabel(file_list_frame, text=t("no_files_selected"), text_color="gray").pack(anchor="w", padx=5, pady=2)
                return
                
            def make_delete_cmd(path_to_delete):
                def _delete():
                    try:
                        if os.path.exists(path_to_delete):
                            os.remove(path_to_delete)
                    except Exception as e:
                        print(f"Error deleting file: {e}")
                        
                    current_paths = files_var.get().split("|")
                    if path_to_delete in current_paths:
                        current_paths.remove(path_to_delete)
                        
                    new_val = "|".join([p for p in current_paths if p])
                    files_var.set(new_val)
                    update_files_display(new_val.split("|") if new_val else [])
                return _delete
                
            for p in valid_paths:
                item_frame = ctk.CTkFrame(file_list_frame, fg_color="transparent")
                item_frame.pack(fill="x", pady=2)
                
                lbl = ctk.CTkLabel(item_frame, text=os.path.basename(p), anchor="w")
                lbl.pack(side="left", fill="x", expand=True, padx=5)
                
                del_btn = ctk.CTkButton(item_frame, text="X", width=20, fg_color="transparent", text_color="#ff5555", hover_color="#8b0000", command=make_delete_cmd(p))
                del_btn.pack(side="right", padx=5)
                
        update_files_display(existing_atts)

        def select_file():
            paths = filedialog.askopenfilenames()
            if paths:
                email_val = email_var.get().strip()
                profile_name = comp_var.get().strip() or email_val
                safe_profile_name = "".join(c for c in profile_name if c.isalnum() or c in (" ", "-", "_", ".", "@")).strip()
                if not safe_profile_name:
                    safe_profile_name = f"profile_{int(time.time())}"
                    
                target_dir = os.path.join(config.ATTACHMENTS_DIR, safe_profile_name)
                os.makedirs(target_dir, exist_ok=True)
                
                # Get currently existing files
                current_files = [p for p in files_var.get().split("|") if p]
                
                prog_lbl.pack(pady=2)
                prog_bar.pack(fill="x", padx=20, pady=5)
                save_btn_bottom.configure(state="disabled")
                popup.update_idletasks()
                
                added_paths = []
                total = len(paths)
                
                for i, p in enumerate(paths):
                    if os.path.exists(p):
                        try:
                            filename = os.path.basename(p)
                            target_path = os.path.join(target_dir, filename)
                            shutil.copy2(p, target_path)
                            added_paths.append(target_path)
                            
                            prog_bar.set((i + 1) / total)
                            prog_lbl.configure(text=f"Copying files... ({i+1}/{total})")
                            popup.update_idletasks()
                        except Exception as e:
                            messagebox.showerror("Error", f"Failed to copy {filename}: {e}", parent=popup)
                            save_btn_bottom.configure(state="normal")
                            prog_lbl.pack_forget()
                            prog_bar.pack_forget()
                            return
                
                # Merge: replace old entries with same filename, add new ones
                added_basenames = {os.path.basename(p) for p in added_paths}
                kept = [p for p in current_files if os.path.basename(p) not in added_basenames]
                all_paths = kept + added_paths
                files_var.set("|".join(all_paths))
                update_files_display(all_paths)
                save_btn_bottom.configure(state="normal")
                prog_lbl.pack_forget()
                prog_bar.pack_forget()

        sync_state = {"folder": None, "active": False}
        
        def check_folder_sync():
            if not popup.winfo_exists() or not sync_state["active"] or not sync_state["folder"]:
                return
                
            folder = sync_state["folder"]
            if os.path.exists(folder):
                current_files = [os.path.join(folder, f) for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
                current_files_str = "|".join(current_files)
                
                if current_files_str != files_var.get():
                    files_var.set(current_files_str)
                    update_files_display(current_files)
                    
            popup.after(1000, check_folder_sync)

        def open_attachment_folder():
            paths = [resolve_att_path(p) for p in files_var.get().split("|") if p]
            folder_to_open = None
            
            # Try to get folder from first valid attachment
            for p in paths:
                d = os.path.normpath(os.path.dirname(p))
                if d and os.path.isdir(d):
                    folder_to_open = d
                    break
            
            # Fallback: create/use profile folder in ATTACHMENTS_DIR
            if not folder_to_open:
                email_val = email_var.get().strip()
                profile_name = comp_var.get().strip() or email_val
                safe_profile_name = "".join(c for c in profile_name if c.isalnum() or c in (" ", "-", "_", ".", "@")).strip()
                if not safe_profile_name:
                    safe_profile_name = f"profile_{int(time.time())}"
                os.makedirs(config.ATTACHMENTS_DIR, exist_ok=True)
                folder_to_open = os.path.normpath(os.path.join(config.ATTACHMENTS_DIR, safe_profile_name))
                os.makedirs(folder_to_open, exist_ok=True)
            
            if folder_to_open and os.path.isdir(folder_to_open):
                os.startfile(folder_to_open)
                
                if not sync_state["active"]:
                    sync_state["folder"] = folder_to_open
                    sync_state["active"] = True
                    popup.after(1000, check_folder_sync)

        ctk.CTkButton(bottom_bar, text=t("open_folder"), command=open_attachment_folder, width=130, fg_color="gray50", hover_color="gray40").pack(side="left", padx=5)
        ctk.CTkButton(bottom_bar, text=t("add_select_files"), command=select_file, width=150).pack(side="left", padx=5)

    def send_all_mails(self):
        smtp_conf = self.app_config.get("smtp", {})
        if not smtp_conf.get("server") or not smtp_conf.get("port") or not smtp_conf.get("user") or not smtp_conf.get("password"):
            messagebox.showerror(t("error"), t("smtp_error"))
            return
            
        threading.Thread(target=self._mailing_engine_worker, daemon=True).start()

    def _mailing_engine_worker(self):
        smtp_conf = self.app_config.get("smtp", {})
        server = smtp_conf.get("server")
        port = int(smtp_conf.get("port", 587))
        user = smtp_conf.get("user")
        password = smtp_conf.get("password") # This is the encrypted password

        contacts = self.app_config.get("contacts", [])
        
        current_run = {
            "date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "deliveries": []
        }
        self.reports_data.append(current_run)
        
        def log_report(email, subj, status_msg, msg_body="", attachment=""):
            if status_msg == "Skipped (Disabled)":
                return # User requested not to log skipped items
                
            current_run["deliveries"].append({
                "email": email,
                "subject": subj,
                "status": status_msg,
                "message": msg_body,
                "attachment": attachment
            })
            self.save_reports()
        
        
        stats = {"sent": 0, "error": 0, "skipped": 0}
        
        try:
            with smtplib.SMTP(server, port) as smtp:
                smtp.ehlo()
                smtp.starttls()
                try:
                    # Always decrypt the password to memory before usage
                    pwd = crypto.decrypt(password)
                    smtp.login(user, pwd)
                except Exception as e:
                    messagebox.showerror("SMTP Login Error", f"Failed to login to SMTP server: {e}")
                    log_report("SYSTEM", "N/A", f"SMTP Login Error: {e}", "")
                    return # Exit the worker if login fails
                
                delay_seconds = int(self.app_config.get("delay", 2))
                
                for i, contact in enumerate(contacts):
                    c_id = contact.get("id")
                    recipient = contact.get("email", "Unknown")
                    subject = contact.get("subject", "")
                    widgets = self.contact_widgets.get(c_id, {})
                    status_lbl = widgets.get("status_lbl")
                    row_frame = widgets.get("row_frame")
                        
                    if not contact.get("enabled", True):
                        msg = "Skipped (Disabled)"
                        if status_lbl: status_lbl.configure(text=msg, text_color="gray")
                        log_report(recipient, subject, msg, contact.get("message", ""))
                        stats["skipped"] += 1
                        continue
                        
                    if status_lbl: status_lbl.configure(text=t("sending"), text_color="yellow")
                    if row_frame:
                        row_frame.configure(border_width=2, border_color="yellow")
                    
                    legacy_att = contact.get("attachment")
                    atts_list = contact.get("attachments", [legacy_att] if legacy_att else [])
                    atts_list = [resolve_att_path(p) for p in atts_list]
                    att_bases = [os.path.basename(p) for p in atts_list if p]
                    att_summary = ", ".join(att_bases) if att_bases else ""
                    
                    if not recipient or recipient == "Unknown":
                        msg = "Error: No Email"
                        if status_lbl: status_lbl.configure(text=msg, text_color="red")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(recipient, subject, msg, contact.get("message", ""), att_summary)
                        continue
                        
                    missing_files = [p for p in atts_list if p and not os.path.exists(p)]
                    if missing_files:
                        msg = "Error: File(s) Missing"
                        if status_lbl: status_lbl.configure(text=msg, text_color="red")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(recipient, subject, msg, contact.get("message", ""), att_summary)
                        continue
                        
                    # Process Dynamic Template Variables
                    now = datetime.datetime.now()
                    
                    date_map = {"DD/MM/YYYY": "%d/%m/%Y", "MM/DD/YYYY": "%m/%d/%Y", "YYYY-MM-DD": "%Y-%m-%d", "DD.MM.YYYY": "%d.%m.%Y"}
                    time_map = {"24H": "%H:%M", "12H": "%I:%M %p"}
                    
                    date_fmt_key = self.app_config.get("date_format", "DD/MM/YYYY")
                    time_fmt_key = self.app_config.get("time_format", "24H")
                    
                    date_fmt = date_fmt_key if "%" in date_fmt_key else date_map.get(date_fmt_key, "%d/%m/%Y")
                    time_fmt = time_fmt_key if "%" in time_fmt_key else time_map.get(time_fmt_key, "%H:%M")
                    
                    current_date_str = now.strftime(date_fmt)
                    current_time_str = now.strftime(time_fmt)
                    
                    msg_body = contact.get("message", "")
                    global_sig = self.app_config.get("signature", "")
            
                    # Fill predefined template variables
                    replacements = {
                        "{company_name}": contact.get("company", ""),
                        "{email}": recipient,
                        "{email_prefix}": recipient.split("@")[0] if "@" in recipient else recipient,
                        "{date}": current_date_str,
                        "{time}": current_time_str,
                        "{signature}": global_sig
                    }
                    
                    for key, val in replacements.items():
                        msg_body = msg_body.replace(key, val)
                        subject = subject.replace(key, val)
                    
                    msg = EmailMessage()
                    msg['Subject'] = subject
                    msg['From'] = user
                    msg['To'] = recipient
                    msg.set_content(msg_body)
                    
                    file_attach_err = False
                    for attachment_path in atts_list:
                        if attachment_path and os.path.exists(attachment_path):
                            try:
                                with open(attachment_path, 'rb') as f:
                                    file_data = f.read()
                                    file_name = os.path.basename(attachment_path)
                                mime_type, _ = mimetypes.guess_type(file_name)
                                if mime_type is None:
                                    mime_type = 'application/octet-stream'
                                maintype, subtype = mime_type.split('/', 1)
                                msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)
                            except Exception as e:
                                errMsg = f"Error: Attaching {file_name}"
                                if status_lbl: status_lbl.configure(text=errMsg, text_color="red")
                                if row_frame: row_frame.configure(border_width=0)
                                log_report(recipient, subject, errMsg, msg_body, att_summary)
                                file_attach_err = True
                                break
                                
                    if file_attach_err:
                        continue

                    try:
                        smtp.send_message(msg)
                        if status_lbl: status_lbl.configure(text="Sent", text_color="green")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(recipient, subject, "Sent", msg_body, att_summary)
                        stats["sent"] += 1
                    except Exception as e:
                        if status_lbl: status_lbl.configure(text="Error: Sending", text_color="red")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(recipient, subject, "Error: Sending", msg_body, att_summary)
                        stats["error"] += 1
                        
                    # Anti-spam Delay (skip if no more enabled contacts remain)
                    remaining_enabled = any(c.get("enabled", True) for c in contacts[i+1:])
                    if remaining_enabled and delay_seconds > 0:
                        import time
                        time.sleep(delay_seconds)

            if self.app_config.get("notifications", True):
                try:
                    from win10toast import ToastNotifier
                    toaster = ToastNotifier()
                    toaster.show_toast(
                        t("notif_title"),
                        t("notif_body", sent=stats['sent'], errors=stats['error'], skipped=stats['skipped']),
                        icon_path=resource_path(os.path.join("assets", "logo.ico")),
                        duration=5,
                        threaded=True
                    )
                except Exception as ex:
                    print(f"Failed to show Windows notification: {ex}")

        except Exception as e:
            messagebox.showerror("SMTP Connection Error", f"Failed to connect: {e}")
            log_report("SYSTEM", "N/A", f"SMTP Error: {e}", "")

if __name__ == "__main__":
    config.reload_paths()
    app = MailAttackerApp()
    if "--view" in sys.argv:
        idx = sys.argv.index("--view")
        if idx + 1 < len(sys.argv):
            view = sys.argv[idx + 1]
            if view == "settings":
                app.show_settings_view()
    app.mainloop()
