import os
import json
import shutil
import threading
import smtplib
from email.message import EmailMessage
import mimetypes
import customtkinter as ctk
from tkinter import filedialog, messagebox

import config

ctk.set_appearance_mode(config.DEFAULT_THEME)
ctk.set_default_color_theme(config.ACCENT_COLOR)

class MailAttackerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title(f"{config.APP_NAME} {config.APP_VERSION}")
        self.geometry(f"{config.WINDOW_WIDTH}x{config.WINDOW_HEIGHT}")
        self.resizable(config.RESIZABLE, config.RESIZABLE)
        
        # Internal Data
        self.app_config = {"smtp": {}, "contacts": []}
        self.contact_widgets = {} # For UI updates during sending without polluting JSON
        self.load_config()

        # Ensure attachments dir exists
        if not os.path.exists(config.ATTACHMENTS_DIR):
            os.makedirs(config.ATTACHMENTS_DIR)

        # Configure Grid
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text=f"{config.APP_NAME}\n{config.APP_VERSION}", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Sidebar navigation buttons
        self.contacts_button = ctk.CTkButton(self.sidebar_frame, text="Contacts", command=self.show_contacts_view)
        self.contacts_button.grid(row=1, column=0, padx=20, pady=10)

        self.settings_button = ctk.CTkButton(self.sidebar_frame, text="Settings", command=self.show_settings_view)
        self.settings_button.grid(row=2, column=0, padx=20, pady=10)
        
        self.send_all_button = ctk.CTkButton(self.sidebar_frame, text="Send All", command=self.send_all_mails, fg_color="green", hover_color="darkgreen")
        self.send_all_button.grid(row=5, column=0, padx=20, pady=(10, 20))

        # Main Frame - Contacts View
        self.contacts_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        # Add Contact Section at the VERY TOP
        self.add_contact_frame = ctk.CTkFrame(self.contacts_frame, fg_color="transparent")
        self.add_contact_frame.pack(fill="x", padx=config.PAD_X, pady=(20, 5))
        
        self.add_btn_top = ctk.CTkButton(self.add_contact_frame, text="+ Add New Contact", fg_color="#1f538d", command=lambda: self.open_contact_popup(None))
        self.add_btn_top.pack(side="left")

        # Contacts List Scrollable Frame
        self.contacts_scrollable_frame = ctk.CTkScrollableFrame(self.contacts_frame, label_text="Contact List")
        self.contacts_scrollable_frame.pack(fill="both", expand=True, padx=config.PAD_X, pady=(5, 20))

        # Main Frame - Settings View
        self.settings_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        
        self.settings_title = ctk.CTkLabel(self.settings_frame, text="SMTP Settings", font=ctk.CTkFont(size=20, weight="bold"))
        self.settings_title.pack(padx=config.PAD_X, pady=20, anchor="w")
        
        self.smtp_server_var = ctk.StringVar(value=self.app_config["smtp"].get("server", ""))
        self.smtp_port_var = ctk.StringVar(value=self.app_config["smtp"].get("port", ""))
        self.smtp_user_var = ctk.StringVar(value=self.app_config["smtp"].get("user", ""))
        self.smtp_pass_var = ctk.StringVar(value=self.app_config["smtp"].get("password", ""))

        form_frame = ctk.CTkFrame(self.settings_frame)
        form_frame.pack(fill="x", padx=config.PAD_X, pady=10)
        
        # Server Label & Entry
        ctk.CTkLabel(form_frame, text="SMTP Server Account (e.g. smtp.gmail.com)", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=self.smtp_server_var).pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        
        # Port Label & Entry
        ctk.CTkLabel(form_frame, text="SMTP Port (e.g. 587)", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=self.smtp_port_var).pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        
        # Email Label & Entry
        ctk.CTkLabel(form_frame, text="Sender Email Address", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=self.smtp_user_var).pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        
        # Password Label & Entry
        ctk.CTkLabel(form_frame, text="Sender Password / App Password", anchor="w").pack(fill="x", padx=config.PAD_X, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=self.smtp_pass_var, show="*").pack(fill="x", padx=config.PAD_X, pady=config.ENTRY_PAD_Y)
        
        ctk.CTkButton(form_frame, text="Save Settings", command=self.save_settings).pack(pady=20)

        # Start with Contacts view
        self.show_contacts_view()

    def load_config(self):
        if os.path.exists(config.CONFIG_FILE):
            try:
                with open(config.CONFIG_FILE, "r") as f:
                    self.app_config = json.load(f)
            except Exception as e:
                print(f"Error loading config: {e}")

        if "smtp" not in self.app_config:
            self.app_config["smtp"] = {}
        if "contacts" not in self.app_config:
            self.app_config["contacts"] = []

        for i, c in enumerate(self.app_config["contacts"]):
            if "id" not in c:
                c["id"] = i
            if "_status_lbl_ref" in c:
                del c["_status_lbl_ref"]

    def save_config(self):
        with open(config.CONFIG_FILE, "w") as f:
            json.dump(self.app_config, f, indent=4)

    def show_contacts_view(self):
        self.settings_frame.grid_forget()
        self.contacts_frame.grid(row=0, column=1, sticky="nsew")
        self.refresh_contacts_list()

    def show_settings_view(self):
        self.contacts_frame.grid_forget()
        self.settings_frame.grid(row=0, column=1, sticky="nsew")

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

        info_text = f"{contact.get('company', '')} | {contact.get('email', '')} | {contact.get('subject', '')}"
        
        lbl = ctk.CTkLabel(row_frame, text=info_text, anchor="w")
        lbl.pack(side="left", padx=10, expand=True, fill="x")

        status_lbl = ctk.CTkLabel(row_frame, text="Ready", width=120)
        status_lbl.pack(side="left", padx=10)
        
        self.contact_widgets[contact.get("id")] = {"status_lbl": status_lbl}

        edit_btn = ctk.CTkButton(row_frame, text="Edit", width=60, command=lambda c=contact: self.open_contact_popup(c))
        edit_btn.pack(side="left", padx=5)

        del_btn = ctk.CTkButton(row_frame, text="Del", width=60, fg_color="red", hover_color="darkred", command=lambda c=contact: self.delete_contact(c))
        del_btn.pack(side="left", padx=5)

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

        # Company Name
        ctk.CTkLabel(form_frame, text="Company Name", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=comp_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        # Recipient Email
        ctk.CTkLabel(form_frame, text="Recipient Email", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=email_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        # Subject
        ctk.CTkLabel(form_frame, text="Subject", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        ctk.CTkEntry(form_frame, textvariable=subj_var).pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)

        # Message Body
        ctk.CTkLabel(form_frame, text="Message Body", anchor="w").pack(fill="x", padx=5, pady=config.LABEL_PAD_Y)
        msg_textbox = ctk.CTkTextbox(form_frame, height=120)
        msg_textbox.pack(fill="x", padx=5, pady=config.ENTRY_PAD_Y)
        if contact:
            msg_textbox.insert("1.0", contact.get("message", ""))

        # Attachment file
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
                target_path = os.path.join(config.ATTACHMENTS_DIR, filename)
                try:
                    idx = 1
                    while os.path.exists(target_path) and os.path.abspath(target_path) != os.path.abspath(new_file_path):
                        name, ext = os.path.splitext(filename)
                        target_path = os.path.join(config.ATTACHMENTS_DIR, f"{name}_{idx}{ext}")
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
                    "attachment": attachment_rel
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
        
        try:
            with smtplib.SMTP(server, port) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.login(user, password)
                
                for contact in contacts:
                    c_id = contact.get("id")
                    widgets = self.contact_widgets.get(c_id, {})
                    status_lbl = widgets.get("status_lbl")
                        
                    if status_lbl:
                        status_lbl.configure(text="Sending...", text_color="yellow")
                    
                    email_to = contact.get("email")
                    if not email_to:
                        if status_lbl: status_lbl.configure(text="Error: No Email", text_color="red")
                        continue
                        
                    attachment_path = contact.get("attachment")
                    
                    if attachment_path and not os.path.exists(attachment_path):
                        if status_lbl: status_lbl.configure(text="Error: File Missing", text_color="red")
                        continue
                        
                    msg = EmailMessage()
                    msg['Subject'] = contact.get("subject", "")
                    msg['From'] = user
                    msg['To'] = email_to
                    msg.set_content(contact.get("message", ""))
                    
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
                            if status_lbl: status_lbl.configure(text="Error: Attaching file", text_color="red")
                            continue

                    try:
                        smtp.send_message(msg)
                        if status_lbl: status_lbl.configure(text="Sent", text_color="green")
                    except Exception as e:
                        if status_lbl: status_lbl.configure(text="Error: Sending", text_color="red")

        except Exception as e:
            messagebox.showerror("SMTP Connection Error", f"Failed to connect: {e}")

if __name__ == "__main__":
    app = MailAttackerApp()
    app.mainloop()
