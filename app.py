import os
import sys
import json
import shutil
import threading
import smtplib
from email.message import EmailMessage
import mimetypes
import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox

import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("eren.mailattacker.v1")

import config

ctk.set_appearance_mode(config.DEFAULT_THEME)
ctk.set_default_color_theme(config.ACCENT_COLOR)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

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

        if not os.path.exists(config.ATTACHMENTS_DIR):
            os.makedirs(config.ATTACHMENTS_DIR)

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text=f"{config.APP_NAME}\n{config.APP_VERSION}", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.contacts_button = ctk.CTkButton(self.sidebar_frame, text="Contacts", command=self.show_contacts_view)
        self.contacts_button.grid(row=1, column=0, padx=20, pady=10)

        self.settings_button = ctk.CTkButton(self.sidebar_frame, text="Settings", command=self.show_settings_view)
        self.settings_button.grid(row=2, column=0, padx=20, pady=10)
        
        self.reports_button = ctk.CTkButton(self.sidebar_frame, text="Reports", command=self.show_reports_view)
        self.reports_button.grid(row=3, column=0, padx=20, pady=10)
        
        self.send_all_button = ctk.CTkButton(self.sidebar_frame, text="Send All", command=self.send_all_mails, fg_color="green", hover_color="darkgreen")
        self.send_all_button.grid(row=6, column=0, padx=20, pady=(10, 20))

        # --- Contacts Frame ---
        self.contacts_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        self.add_contact_frame = ctk.CTkFrame(self.contacts_frame, fg_color="transparent")
        self.add_contact_frame.pack(fill="x", padx=config.PAD_X, pady=(20, 5))
        
        self.add_btn_top = ctk.CTkButton(self.add_contact_frame, text="+ Add New Contact", fg_color="#1f538d", command=lambda: self.open_contact_popup(None))
        self.add_btn_top.pack(side="left")
        
        self.toggle_all_var = ctk.BooleanVar(value=True)
        self.toggle_all_btn = ctk.CTkButton(self.add_contact_frame, text="Select/Deselect All", fg_color="gray30", hover_color="gray40", command=self.toggle_all_contacts)
        self.toggle_all_btn.pack(side="right")

        self.contacts_scrollable_frame = ctk.CTkScrollableFrame(self.contacts_frame, label_text="Contact List")
        self.contacts_scrollable_frame.pack(fill="both", expand=True, padx=config.PAD_X, pady=(5, 20))
        
        # Context Menu for Right Click
        import tkinter as tk
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Toggle Selected", command=self.toggle_selected_contacts)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Select All", command=lambda: self.set_all_contacts(True))
        self.context_menu.add_command(label="Deselect All", command=lambda: self.set_all_contacts(False))
        
        self.contacts_scrollable_frame.bind("<Button-3>", self.show_context_menu)

        # --- Settings Frame ---
        self.settings_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        self.settings_title = ctk.CTkLabel(self.settings_frame, text="SMTP Settings", font=ctk.CTkFont(size=20, weight="bold"))
        self.settings_title.pack(padx=config.PAD_X, pady=20, anchor="w")
        
        self.smtp_server_var = ctk.StringVar(value=self.app_config["smtp"].get("server", ""))
        self.smtp_port_var = ctk.StringVar(value=self.app_config["smtp"].get("port", ""))
        self.smtp_user_var = ctk.StringVar(value=self.app_config["smtp"].get("user", ""))
        self.smtp_pass_var = ctk.StringVar(value=self.app_config["smtp"].get("password", ""))

        form_frame = ctk.CTkFrame(self.settings_frame)
        form_frame.pack(fill="x", padx=config.PAD_X, pady=10)
        
        # Server Input
        ctk.CTkLabel(form_frame, text="SMTP Server Account (e.g. smtp.gmail.com)", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        server_options = ["smtp.gmail.com", "smtp.office365.com", "smtp.mail.yahoo.com"]
        self.smtp_server_cb = ctk.CTkComboBox(form_frame, values=server_options, variable=self.smtp_server_var)
        self.smtp_server_cb.pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        
        # Port Input with validation
        vcmd_port = (self.register(self.validate_port), '%P')
        ctk.CTkLabel(form_frame, text="SMTP Port (e.g. 587)", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=self.smtp_port_var, validate='key', validatecommand=vcmd_port).pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        
        # Email Input with Validation
        ctk.CTkLabel(form_frame, text="Sender Email Address", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        self.email_entry = ctk.CTkEntry(form_frame, textvariable=self.smtp_user_var, placeholder_text="example@domain.com")
        self.email_entry.pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        self.email_entry.bind("<FocusOut>", self.validate_email_format)
        
        # Password Input with Toggle
        ctk.CTkLabel(form_frame, text="Sender Password / App Password", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        pass_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        pass_frame.pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        
        self.pass_entry = ctk.CTkEntry(pass_frame, textvariable=self.smtp_pass_var, show="*")
        self.pass_entry.pack(side="left", fill="x", expand=True)
        
        self.show_pass_var = ctk.BooleanVar(value=False)
        self.show_pass_btn = ctk.CTkCheckBox(pass_frame, text="👁", variable=self.show_pass_var, command=self.toggle_password_visibility, width=30)
        self.show_pass_btn.pack(side="right", padx=(5, 0))
        
        ctk.CTkButton(form_frame, text="Save Settings", command=self.save_settings).pack(pady=20)
        
        # --- Reports Frame ---
        self.reports_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        report_top_frame = ctk.CTkFrame(self.reports_frame, fg_color="transparent")
        report_top_frame.pack(fill="x", padx=config.PAD_X, pady=20)
        
        self.reports_title = ctk.CTkLabel(report_top_frame, text="Send Reports", font=ctk.CTkFont(size=20, weight="bold"))
        self.reports_title.pack(side="left")
        
        self.clear_reports_btn = ctk.CTkButton(report_top_frame, text="Clear Reports", fg_color="red", hover_color="darkred", command=self.clear_reports, width=100)
        self.clear_reports_btn.pack(side="right")
        
        self.reports_scrollable_frame = ctk.CTkScrollableFrame(self.reports_frame, label_text="Recent Logs")
        self.reports_scrollable_frame.pack(fill="both", expand=True, padx=config.PAD_X, pady=(5, 20))

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
        self.contacts_frame.grid(row=0, column=1, sticky="nsew")
        self.refresh_contacts_list()

    def show_settings_view(self):
        self.contacts_frame.grid_forget()
        self.reports_frame.grid_forget()
        self.settings_frame.grid(row=0, column=1, sticky="nsew")
        
    def show_reports_view(self):
        self.contacts_frame.grid_forget()
        self.settings_frame.grid_forget()
        self.reports_frame.grid(row=0, column=1, sticky="nsew")
        self.refresh_reports_list()

    def clear_reports(self):
        if messagebox.askyesno("Clear Reports", "Are you sure you want to clear all report logs?"):
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

            summary_text = f"Send Batch - {run.get('date', 'Unknown Time')} ({len(deliveries)} emails)"
            btn = ctk.CTkButton(run_frame, text=summary_text, anchor="w", fg_color="#1f538d",
                                command=lambda sf=subframe: self.toggle_report_group(sf))
            btn.pack(fill="x")
            
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
        self.report_popup.title("Message Details")
        self.report_popup.geometry("500x450")
        self.report_popup.attributes("-topmost", True)
        self.report_popup.transient(self)
        
        self.report_popup.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() // 2) - (500 // 2)
        y = self.winfo_rooty() + (self.winfo_height() // 2) - (450 // 2)
        self.report_popup.geometry(f"+{x}+{y}")
        
        info_frame = ctk.CTkFrame(self.report_popup, fg_color="transparent")
        info_frame.pack(fill="x", padx=20, pady=(20, 10))
        
        ctk.CTkLabel(info_frame, text=f"To: {rep_data.get('email', '')}", font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x")
        ctk.CTkLabel(info_frame, text=f"Subject: {rep_data.get('subject', '')}", font=ctk.CTkFont(weight="bold"), anchor="w").pack(fill="x", pady=(5,0))
        
        att_name = rep_data.get('attachment', '') or "None"
        ctk.CTkLabel(info_frame, text=f"Attachment: {att_name}", font=ctk.CTkFont(weight="bold"), anchor="w", text_color="lightgray").pack(fill="x", pady=(5,0))
        
        msg_frame = ctk.CTkFrame(self.report_popup)
        msg_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        textbox = ctk.CTkTextbox(msg_frame)
        textbox.pack(fill="both", expand=True, padx=2, pady=2)
        textbox.insert("1.0", rep_data.get("message", "No message content recorded."))
        textbox.configure(state="disabled")

    def save_settings(self):
        self.app_config["smtp"] = {
            "server": self.smtp_server_var.get().strip(),
            "port": self.smtp_port_var.get().strip(),
            "user": self.smtp_user_var.get().strip(),
            "password": self.smtp_pass_var.get().strip()
        }
        self.save_config()
        messagebox.showinfo("Success", "SMTP settings saved.")

    def refresh_contacts_list(self):
        for widget in self.contacts_scrollable_frame.winfo_children():
            widget.destroy()
        
        self.contact_widgets.clear()
        for c in self.app_config["contacts"]:
            self.create_contact_row(c)

    def create_contact_row(self, contact):
        row_frame = ctk.CTkFrame(self.contacts_scrollable_frame)
        row_frame.pack(fill="x", pady=5, padx=5)
        
        row_frame.contact_id = contact.get("id")
        
        row_frame.grid_columnconfigure(1, weight=0)
        row_frame.grid_columnconfigure(2, weight=1, uniform="col")
        row_frame.grid_columnconfigure(3, weight=2, uniform="col")
        row_frame.grid_columnconfigure(4, weight=2, uniform="col")
        row_frame.grid_columnconfigure(5, weight=0)
        row_frame.grid_columnconfigure(6, weight=0)
        row_frame.grid_columnconfigure(7, weight=0)
        
        drag_handle = ctk.CTkLabel(row_frame, text="☰", cursor="fleur", width=30)
        drag_handle.grid(row=0, column=0, padx=(10, 5), pady=10)

        is_enabled = ctk.BooleanVar(value=contact.get('enabled', True))
        
        enable_cb = ctk.CTkCheckBox(row_frame, text="", variable=is_enabled, width=20)
        enable_cb.grid(row=0, column=1, padx=5)

        def truncate(text, max_len=20):
            return text if len(text) <= max_len else text[:max_len-3] + "..."
            
        font_style = ctk.CTkFont(overstrike=not is_enabled.get(), slant="italic" if not is_enabled.get() else "roman")
        text_color = "gray50" if not is_enabled.get() else ["gray10", "#DCE4EE"]

        company_lbl = ctk.CTkLabel(row_frame, text=truncate(contact.get('company', '')), anchor="w", cursor="fleur", text_color=text_color, font=font_style)
        company_lbl.grid(row=0, column=2, sticky="w", padx=5)
        
        email_lbl = ctk.CTkLabel(row_frame, text=truncate(contact.get('email', ''), 25), anchor="w", cursor="fleur", text_color=text_color, font=font_style)
        email_lbl.grid(row=0, column=3, sticky="w", padx=5)
        
        subj_lbl = ctk.CTkLabel(row_frame, text=truncate(contact.get('subject', ''), 25), anchor="w", cursor="fleur", text_color=text_color, font=font_style)
        subj_lbl.grid(row=0, column=4, sticky="w", padx=5)

        status_lbl = ctk.CTkLabel(row_frame, text="Ready", width=80)
        status_lbl.grid(row=0, column=5, padx=10)
        
        self.contact_widgets[contact.get("id")] = {
            "status_lbl": status_lbl,
            "is_enabled_var": is_enabled,
            "row_frame": row_frame
        }
        
        edit_btn = ctk.CTkButton(row_frame, text="Edit", width=60, command=lambda c=contact: self.open_contact_popup(c))
        edit_btn.grid(row=0, column=6, padx=5)

        del_btn = ctk.CTkButton(row_frame, text="Del", width=60, fg_color="red", hover_color="darkred", command=lambda c=contact: self.delete_contact(c))
        del_btn.grid(row=0, column=7, padx=(5, 10))
        
        def toggle_contact_state():
            enabled = is_enabled.get()
            new_f_style = ctk.CTkFont(overstrike=not enabled, slant="italic" if not enabled else "roman")
            new_t_color = "gray50" if not enabled else ["gray10", "#DCE4EE"]
            company_lbl.configure(text_color=new_t_color, font=new_f_style)
            email_lbl.configure(text_color=new_t_color, font=new_f_style)
            subj_lbl.configure(text_color=new_t_color, font=new_f_style)
            
            c_id = contact.get("id")
            for c in self.app_config["contacts"]:
                if c.get("id") == c_id:
                    c["enabled"] = enabled
                    self.save_config()
                    break

        enable_cb.configure(command=toggle_contact_state)
        self.contact_widgets[contact.get("id")]["toggle_cmd"] = toggle_contact_state
        
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
                msg = f"Are you sure you want to delete {len(to_delete)} contacts?" if len(to_delete) > 1 else "Are you sure you want to delete this contact?"
                if messagebox.askyesno("Delete", msg):
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
        
        info_text = f"Moving {len(self.dragged_rows)} Contact(s)..."
        
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
        if messagebox.askyesno("Delete", f"Are you sure you want to delete {contact_to_del.get('email')}?"):
            self.app_config["contacts"] = [c for c in self.app_config["contacts"] if c.get("id") != contact_to_del.get("id")]
            self.save_config()
            self.refresh_contacts_list()

    def open_contact_popup(self, contact=None):
        popup = ctk.CTkToplevel(self)
        popup.title("Add New Contact" if contact is None else "Edit Contact")
        popup.geometry("600x650")
        popup.transient(self)
        popup.grab_set()

        icon_path = resource_path(os.path.join("assets", "logo.ico"))
        if os.path.exists(icon_path):
            popup.iconbitmap(icon_path)

        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 600) // 2
        y = self.winfo_y() + (self.winfo_height() - 650) // 2
        popup.geometry(f"+{x}+{y}")

        comp_var = ctk.StringVar(value=contact.get("company", "") if contact else "")
        email_var = ctk.StringVar(value=contact.get("email", "") if contact else "")
        subj_var = ctk.StringVar(value=contact.get("subject", "") if contact else "")
        file_var = ctk.StringVar(value=contact.get("attachment", "") if contact else "")

        form_frame = ctk.CTkScrollableFrame(popup, fg_color="transparent")
        form_frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(form_frame, text="Company Name", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=comp_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        ctk.CTkLabel(form_frame, text="Recipient Email", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=email_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        ctk.CTkLabel(form_frame, text="Subject", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=subj_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        ctk.CTkLabel(form_frame, text="Message Body", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        msg_textbox = ctk.CTkTextbox(form_frame, height=120)
        msg_textbox.pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)
        if contact:
            msg_textbox.insert("1.0", contact.get("message", ""))

        ctk.CTkLabel(form_frame, text="Attachment File", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        
        file_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        file_frame.pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)
        
        file_lbl = ctk.CTkLabel(file_frame, text=os.path.basename(file_var.get()) if file_var.get() else "No file selected", anchor="w")
        file_lbl.pack(side="left", fill="x", expand=True)

        def select_file():
            path = filedialog.askopenfilename()
            if path:
                file_var.set(path)
                file_lbl.configure(text=os.path.basename(path))

        ctk.CTkButton(file_frame, text="Select File", command=select_file, width=120).pack(side="right", padx=5)

        def save_changes():
            email_val = email_var.get().strip()
            if not email_val:
                messagebox.showerror("Error", "Recipient Email is required", parent=popup)
                return

            new_file_path = file_var.get()
            attachment_rel = contact.get("attachment", "") if contact else ""

            if new_file_path and new_file_path != attachment_rel and os.path.exists(new_file_path):
                filename = os.path.basename(new_file_path)
                
                profile_name = comp_var.get().strip() or email_val
                safe_profile_name = "".join(c for c in profile_name if c.isalnum() or c in (" ", "-", "_", ".", "@")).strip()
                if not safe_profile_name:
                    safe_profile_name = "profile"
                    
                target_dir = os.path.join(config.ATTACHMENTS_DIR, safe_profile_name)
                
                try:
                    if not os.path.exists(target_dir):
                        os.makedirs(target_dir)
                        
                    target_path = os.path.join(target_dir, filename)
                    
                    idx = 1
                    while os.path.exists(target_path) and os.path.abspath(target_path) != os.path.abspath(new_file_path):
                        name, ext = os.path.splitext(filename)
                        target_path = os.path.join(target_dir, f"{name}_{idx}{ext}")
                        idx += 1
                    if os.path.abspath(target_path) != os.path.abspath(new_file_path):
                        shutil.copy2(new_file_path, target_path)
                    attachment_rel = target_path
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy file: {e}", parent=popup)
                    return
            elif not new_file_path:
                attachment_rel = ""

            if contact:
                c_id = contact.get("id")
                for c in self.app_config["contacts"]:
                    if c.get("id") == c_id:
                        c["company"] = comp_var.get().strip()
                        c["email"] = email_val
                        c["subject"] = subj_var.get().strip()
                        c["message"] = msg_textbox.get("1.0", "end-1c").strip()
                        c["attachment"] = attachment_rel
                        break
            else:
                new_id = 0 if not self.app_config["contacts"] else max(int(c.get("id", 0)) for c in self.app_config["contacts"]) + 1
                new_contact = {
                    "id": new_id,
                    "company": comp_var.get().strip(),
                    "email": email_val,
                    "subject": subj_var.get().strip(),
                    "message": msg_textbox.get("1.0", "end-1c").strip(),
                    "attachment": attachment_rel,
                    "enabled": True
                }
                self.app_config["contacts"].insert(0, new_contact)

            self.save_config()
            self.refresh_contacts_list()
            popup.destroy()

        ctk.CTkButton(popup, text="Save Contact List" if contact else "Add to Contact List", command=save_changes, fg_color="#1f538d").pack(pady=20)

    def send_all_mails(self):
        smtp_conf = self.app_config.get("smtp", {})
        if not smtp_conf.get("server") or not smtp_conf.get("port") or not smtp_conf.get("user") or not smtp_conf.get("password"):
            messagebox.showerror("SMTP Error", "Please configure SMTP settings first.")
            return
            
        threading.Thread(target=self._mailing_engine_worker, daemon=True).start()

    def _mailing_engine_worker(self):
        smtp_conf = self.app_config.get("smtp", {})
        server = smtp_conf.get("server")
        port = int(smtp_conf.get("port", 587))
        user = smtp_conf.get("user")
        password = smtp_conf.get("password")

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
        
        try:
            with smtplib.SMTP(server, port) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.login(user, password)
                
                for contact in contacts:
                    c_id = contact.get("id")
                    email_to = contact.get("email", "Unknown")
                    subject = contact.get("subject", "")
                    widgets = self.contact_widgets.get(c_id, {})
                    status_lbl = widgets.get("status_lbl")
                    row_frame = widgets.get("row_frame")
                        
                    if not contact.get("enabled", True):
                        msg = "Skipped (Disabled)"
                        if status_lbl: status_lbl.configure(text=msg, text_color="gray")
                        log_report(email_to, subject, msg, contact.get("message", ""))
                        continue
                        
                    if status_lbl:
                        status_lbl.configure(text="Sending...", text_color="yellow")
                    if row_frame:
                        row_frame.configure(border_width=2, border_color="yellow")
                    
                    attachment_path = contact.get("attachment")
                    att_base = os.path.basename(attachment_path) if attachment_path else ""
                    
                    if not email_to or email_to == "Unknown":
                        msg = "Error: No Email"
                        if status_lbl: status_lbl.configure(text=msg, text_color="red")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(email_to, subject, msg, contact.get("message", ""), att_base)
                        continue
                        
                    if attachment_path and not os.path.exists(attachment_path):
                        msg = "Error: File Missing"
                        if status_lbl: status_lbl.configure(text=msg, text_color="red")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(email_to, subject, msg, contact.get("message", ""), att_base)
                        continue
                        
                    msg = EmailMessage()
                    msg['Subject'] = subject
                    msg['From'] = user
                    msg['To'] = email_to
                    msg_body = contact.get("message", "")
                    msg.set_content(msg_body)
                    
                    file_attach_err = False
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
                            errMsg = "Error: Attaching file"
                            if status_lbl: status_lbl.configure(text=errMsg, text_color="red")
                            if row_frame: row_frame.configure(border_width=0)
                            log_report(email_to, subject, errMsg, msg_body, att_base)
                            file_attach_err = True
                            
                    if file_attach_err:
                        continue

                    try:
                        smtp.send_message(msg)
                        if status_lbl: status_lbl.configure(text="Sent", text_color="green")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(email_to, subject, "Sent", msg_body, att_base)
                    except Exception as e:
                        if status_lbl: status_lbl.configure(text="Error: Sending", text_color="red")
                        if row_frame: row_frame.configure(border_width=0)
                        log_report(email_to, subject, "Error: Sending", msg_body, att_base)

        except Exception as e:
            messagebox.showerror("SMTP Connection Error", f"Failed to connect: {e}")
            log_report("SYSTEM", "N/A", f"SMTP Error: {e}", "")

if __name__ == "__main__":
    app = MailAttackerApp()
    app.mainloop()
