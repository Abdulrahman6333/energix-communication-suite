import os
import sys
import time
import mimetypes
import smtplib
import json

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog
from tkinter import ttk

import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from email.message import EmailMessage

from PIL import Image, ImageDraw, ImageFont, ImageTk
import requests
import arabic_reshaper
from bidi.algorithm import get_display

# ================== THEME / COLORS ==================
BG_MAIN   = "#101318"
BG_PANEL  = "#181C24"
FG_TEXT   = "#FFFFFF"
ACCENT    = "#00A2FF"
BTN_BG    = ACCENT
BTN_FG    = "#FFFFFF"
ENTRY_BG  = "#1F2430"
ENTRY_FG  = "#FFFFFF"

APP_TITLE    = "Energix Suite v4"
APP_SUBTITLE = "Python Communication Automation Suite"

CERT_DIR          = "certificates"
FONTS_DIR         = "fonts"
CONFIG_FILE       = "settings.json"
TEMPLATES_WA_DIR  = "templates_whatsapp"
TEMPLATES_EM_DIR  = "templates_email"

DEFAULT_SMTP_SERVER   = "smtp.gmail.com"
DEFAULT_SMTP_PORT     = 465
DEFAULT_EMAIL_SUBJECT = "Energix Academy – Certificate"

# ============ FONTS FOR CERTIFICATES ============
FONT_URLS = {
    "Cairo-Regular.ttf":   "https://github.com/google/fonts/raw/main/ofl/cairo/Cairo-Regular.ttf",
    "Amiri-Regular.ttf":   "https://github.com/aliftype/amiri-font/raw/master/ttf/Amiri-Regular.ttf",
    "Tajawal-Regular.ttf": "https://github.com/google/fonts/raw/main/ofl/tajawal/Tajawal-Regular.ttf",
}
PREFERRED_FONTS = ["Cairo-Regular.ttf", "Amiri-Regular.ttf", "Tajawal-Regular.ttf"]


def log(msg: str):
    print(msg)

# ================== PATH HELPERS ==================
def resource_path(relative_path: str) -> str:
    # Works for dev + PyInstaller onefile
    base_path = getattr(sys, '_MEIPASS', os.path.abspath('.'))
    return os.path.join(base_path, relative_path)



# ================== PHONE NORMALIZATION ==================
def normalize_egypt_number(raw: str) -> str:
    """
    Normalize Egyptian phone numbers to WhatsApp format:
    Always returns something like: 2010XXXXXXXX
    Accepts: +2010..., 2010..., 010..., 10..., etc.
    """
    if not raw:
        return ""
    raw = str(raw).strip()
    for ch in [" ", "-", "(", ")", "\u200f"]:
        raw = raw.replace(ch, "")

    if raw.startswith("+20"):
        raw = "20" + raw[3:]
    elif raw.startswith("20"):
        pass
    elif raw.startswith("0"):
        raw = "20" + raw[1:]
    else:
        if len(raw) in (10, 11):
            raw = "20" + raw

    raw = "".join(filter(str.isdigit, raw))
    return raw


# ================== WHATSAPP SENDING ==================
def send_whatsapp_messages(excel_path: str, phone_column: str, message_template: str,
                          attach_pdf: bool = False, pdf_default_path: str = "", pdf_caption: str = ""):
    df = pd.read_excel(excel_path)
    if phone_column not in df.columns:
        raise ValueError(f"Column '{phone_column}' not found in Excel file!")

    options = webdriver.ChromeOptions()
    # OPTIONAL: use real Chrome profile (avoid QR every time)
    # options.add_argument(r"--user-data-dir=C:\Users\YOUR_USER\AppData\Local\Google\Chrome\User Data")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()

    driver.get("https://web.whatsapp.com/")
    messagebox.showinfo(
        "WhatsApp",
        "Scan the QR code (if needed), wait until chats appear, then click OK."
    )

    # Wait until WhatsApp Web is really loaded (pane-side is the chats list)
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.ID, "pane-side"))
        )
    except Exception:
        pass

    def find_input_box(drv):
        # More robust selector list
        candidates = [
            (By.CSS_SELECTOR, "footer div[contenteditable='true'][role='textbox']"),
            (By.XPATH, "//footer//div[@contenteditable='true' and @role='textbox']"),
            (By.XPATH, "//footer//div[@contenteditable='true' and @data-tab]"),
        ]
        last = None
        for by, sel in candidates:
            try:
                el = WebDriverWait(drv, 25).until(
                    EC.element_to_be_clickable((by, sel))
                )
                return el
            except Exception as e:
                last = e
        raise last

    def paste_multiline(drv, input_el, text: str):
        # Best: clipboard paste (preserves newlines)
        try:
            import pyperclip
            pyperclip.copy(text)
            input_el.send_keys(Keys.CONTROL, 'v')
            return
        except Exception:
            pass

        # Fallback: Shift+Enter between lines
        lines = str(text).split("\n")
        for i, line in enumerate(lines):
            # keep empty lines too
            if line:
                input_el.send_keys(line)
            if i != len(lines) - 1:
                input_el.send_keys(Keys.SHIFT, Keys.ENTER)

    def attach_pdf_file(drv, file_path: str, caption: str = ""):
        if not file_path:
            return False
        file_path = os.path.abspath(str(file_path).strip())
        if not os.path.exists(file_path):
            log(f"[SKIP PDF] File not found: {file_path}")
            return False

        wait = WebDriverWait(drv, 35)

        # Open attach menu
        try:
            clip_btn = wait.until(
                EC.element_to_be_clickable((By.XPATH,
                    "//span[@data-icon='clip' or @data-testid='attach-menu-plus']/ancestor::button | //div[@title='Attach']"
                ))
            )
            clip_btn.click()
            time.sleep(0.5)
        except Exception as e:
            log(f"[PDF] Can't open attach menu: {e}")
            return False

        # Click Document explicitly
        doc_xpaths = [
            "//span[contains(@data-icon,'document')]/ancestor::*[self::button or self::div][1]",
            "//button[@aria-label='Document']",
            "//div[@role='button' and .//span[contains(@data-icon,'document')]]",
            "//span[contains(@data-testid,'attach-document')]/ancestor::*[self::button or self::div][1]",
        ]
        for xp in doc_xpaths:
            try:
                btn = WebDriverWait(drv, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
                btn.click()
                break
            except Exception:
                pass

        # Send file path to input
        try:
            file_input = WebDriverWait(drv, 25).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
            )
            file_input.send_keys(file_path)
        except Exception as e:
            log(f"[PDF] Can't find file input: {e}")
            return False

        time.sleep(1.2)

        # Optional caption
        if caption:
            try:
                cap_box = WebDriverWait(drv, 10).until(
                    EC.presence_of_element_located((By.XPATH,
                        "//div[@contenteditable='true' and (@role='textbox' or @data-tab)]"
                    ))
                )
                cap_box.click()
                cap_box.send_keys(caption)
            except Exception:
                pass

        # Click send
        try:
            send_btn = WebDriverWait(drv, 90).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//span[@data-icon='send' or @data-testid='send']/ancestor::button"
                ))
            )
            send_btn.click()
            time.sleep(1.2)
        except Exception as e:
            log(f"[PDF] Can't click send: {e}")
            return False

        return True

    for _, row in df.iterrows():
        raw_number = str(row[phone_column])
        phone = normalize_egypt_number(raw_number)

        if not phone or not phone.isdigit():
            log(f"[SKIP] Invalid number: {raw_number} -> {phone}")
            continue

        chat_url = f"https://web.whatsapp.com/send?phone={phone}"
        driver.get(chat_url)

        try:
            input_box = find_input_box(driver)
            input_box.click()
            time.sleep(0.8)

            input_box.send_keys(Keys.CONTROL, "a")
            input_box.send_keys(Keys.DELETE)
            time.sleep(0.2)

            paste_multiline(driver, input_box, message_template)

            # Try enter
            try:
                input_box.send_keys(Keys.ENTER)
            except Exception:
                pass

            # Fallback: click Send
            try:
                send_btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH,
                        "//span[@data-icon='send' or @data-testid='send']/ancestor::button"
                    ))
                )
                send_btn.click()
            except Exception:
                pass

            log(f"[✓] WhatsApp message sent to: {phone}")
            time.sleep(1.5)

            if attach_pdf and pdf_default_path:
                ok = attach_pdf_file(driver, pdf_default_path, pdf_caption)
                if ok:
                    log(f"[✓] PDF attached to: {phone}")
                time.sleep(1.0)

            time.sleep(2.0)

        except Exception as e:
            log(f"[×] Failed to send WhatsApp to {phone}: {e}")
            time.sleep(2)

    driver.quit()
    log("All WhatsApp messages have been processed.")
# ================== EMAIL SENDING ==================
def attach_certificate(msg: EmailMessage, name: str):
    safe_name = str(name).strip().replace(" ", "_")
    exts = [".pdf", ".png", ".jpg", ".jpeg"]

    for ext in exts:
        path = os.path.join(CERT_DIR, f"certificate_{safe_name}{ext}")
        if os.path.exists(path):
            ctype, encoding = mimetypes.guess_type(path)
            if ctype is None or encoding is not None:
                ctype = "application/octet-stream"
            maintype, subtype = ctype.split("/", 1)
            with open(path, "rb") as f:
                msg.add_attachment(
                    f.read(),
                    maintype=maintype,
                    subtype=subtype,
                    filename=os.path.basename(path),
                )
            log(f"  [+] Attached certificate: {path}")
            return
    log(f"  [!] No certificate found for {name}")


def send_emails_with_certificates(
    excel_path: str,
    name_column: str,
    email_column: str,
    email_body_template: str,
    sender_email: str,
    sender_password: str,
    email_subject: str,
    attach_certs: bool,
    smtp_server: str = DEFAULT_SMTP_SERVER,
    smtp_port: int = DEFAULT_SMTP_PORT,
):
    df = pd.read_excel(excel_path)
    for col in [name_column, email_column]:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in Excel file!")

    server = smtplib.SMTP_SSL(smtp_server, smtp_port)
    server.login(sender_email, sender_password)
    log("SMTP login successful. Starting email sending...")

    for _, row in df.iterrows():
        name = str(row[name_column]).strip()
        email = str(row[email_column]).strip()

        if not email or "@" not in email:
            log(f"[SKIP] Invalid email for {name}: {email}")
            continue

        msg = EmailMessage()
        msg["From"] = sender_email
        msg["To"] = email
        msg["Subject"] = email_subject

        body = email_body_template.replace("{name}", name)
        msg.set_content(body)

        if attach_certs:
            attach_certificate(msg, name)

        try:
            server.send_message(msg)
            log(f"[✓] Email sent to: {name} <{email}>")
        except Exception as e:
            log(f"[×] Failed to send email to {name} <{email}>: {e}")

    server.quit()
    log("All emails have been processed.")


# ================== TEMPLATE SAVE / LOAD ==================
def save_template(text_widget, templates_dir: str, title: str):
    os.makedirs(templates_dir, exist_ok=True)
    content = text_widget.get("1.0", tk.END).strip()
    if not content:
        messagebox.showerror("Error", f"{title} template box is empty!")
        return

    name = simpledialog.askstring("Save Template", f"{title} template name (no extension):")
    if not name:
        return

    path = os.path.join(templates_dir, f"{name}.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

    messagebox.showinfo("Saved", f"{title} template saved:\n{path}")


def load_template(text_widget, templates_dir: str, title: str):
    os.makedirs(templates_dir, exist_ok=True)
    path = filedialog.askopenfilename(
        title=f"Load {title} Template",
        initialdir=templates_dir,
        filetypes=[("Text files", "*.txt")],
    )
    if not path:
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        text_widget.delete("1.0", tk.END)
        text_widget.insert(tk.END, content)
        messagebox.showinfo("Loaded", f"{title} template loaded from:\n{path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load {title} template:\n{e}")


# ================== SETTINGS (EMAIL + TEMPLATES) ==================
def load_settings():
    if not os.path.exists(CONFIG_FILE):
        return
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return

    sender_email_var.set(data.get("sender_email", ""))
    # Security: never restore an email app password from disk.
    sender_password_var.set("")
    email_subject_var.set(data.get("email_subject", DEFAULT_EMAIL_SUBJECT))
    attach_cert_var.set(1 if data.get("attach_certificates", True) else 0)

    wa_txt = data.get("whatsapp_template", "")
    em_txt = data.get("email_template", "")

    if wa_txt:
        msg_wa_text.delete("1.0", tk.END)
        msg_wa_text.insert(tk.END, wa_txt)
    if em_txt:
        msg_email_text.delete("1.0", tk.END)
        msg_email_text.insert(tk.END, em_txt)


def save_settings():
    data = {
        "sender_email": sender_email_var.get().strip(),

        "email_subject": email_subject_var.get().strip() or DEFAULT_EMAIL_SUBJECT,
        "attach_certificates": bool(attach_cert_var.get()),
        "whatsapp_template": msg_wa_text.get("1.0", tk.END).strip(),
        "email_template": msg_email_text.get("1.0", tk.END).strip(),
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        log("[CONFIG] Settings saved.")
    except Exception as e:
        log(f"[CONFIG] Failed to save settings: {e}")


# ================== CERTIFICATES – FONT HELPERS ==================
def ensure_fonts_dir_and_download():
    os.makedirs(FONTS_DIR, exist_ok=True)
    for fname, url in FONT_URLS.items():
        fpath = os.path.join(FONTS_DIR, fname)
        if not os.path.exists(fpath):
            log(f"[FONT] Downloading {fname} ...")
            try:
                r = requests.get(url, timeout=30)
                r.raise_for_status()
                with open(fpath, "wb") as f:
                    f.write(r.content)
                log(f"[FONT] Downloaded {fname}")
            except Exception as e:
                log(f"[FONT] Failed to download {fname}: {e}")
        else:
            log(f"[FONT] Exists: {fname}")


def get_default_font_path():
    for fname in PREFERRED_FONTS:
        fpath = os.path.join(FONTS_DIR, fname)
        if os.path.exists(fpath):
            return fpath
    fonts = [f for f in os.listdir(FONTS_DIR) if f.lower().endswith((".ttf", ".otf"))]
    if fonts:
        return os.path.join(FONTS_DIR, fonts[0])
    raise FileNotFoundError("No font files found in 'fonts' folder.")


def reshape_if_arabic(text: str) -> str:
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


def draw_name_on_template(
    template_path: str,
    name: str,
    font_path: str,
    font_size: int,
    y_ratio: float,
    x_offset: int,
    is_arabic: bool,
):
    img = Image.open(template_path).convert("RGB")
    draw = ImageDraw.Draw(img)
    img_w, img_h = img.size

    text = reshape_if_arabic(name) if is_arabic else name
    font = ImageFont.truetype(font_path, font_size)

    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    center_x = img_w / 2 + x_offset
    x = int(center_x - text_w / 2)
    center_y = img_h * y_ratio
    y = int(center_y - text_h / 2)

    color = (0, 114, 188)
    draw.text((x, y), text, font=font, fill=color)
    return img


def generate_all_certificates(template_path, excel_path, name_column,
                              font_path, font_size, y_ratio, x_offset, is_arabic):
    df = pd.read_excel(excel_path)
    if name_column not in df.columns:
        raise ValueError(f"Column '{name_column}' not found in Excel file!")

    os.makedirs(CERT_DIR, exist_ok=True)
    count = 0

    for _, row in df.iterrows():
        name = str(row[name_column]).strip()
        if not name:
            continue

        img = draw_name_on_template(template_path, name, font_path,
                                    font_size, y_ratio, x_offset, is_arabic)
        safe_name = name.replace(" ", "_")
        out_path = os.path.join(CERT_DIR, f"certificate_{safe_name}.pdf")
        img.save(out_path, "PDF", resolution=300.0)
        log(f"[CERT] Generated: {out_path}")
        count += 1

    messagebox.showinfo("Certificates", f"Generated {count} certificates in '{CERT_DIR}' folder.")


# ================== GUI CALLBACKS – WHATSAPP TAB ==================

def browse_pdf_whatsapp():
    path = filedialog.askopenfilename(
        title="Select Default PDF",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    if path:
        wa_pdf_default_var.set(path)


def browse_excel_whatsapp():
    path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        excel_path_var.set(path)


def start_whatsapp_sending():
    excel_path = excel_path_var.get().strip()
    phone_col  = phone_col_var.get().strip()
    msg        = msg_wa_text.get("1.0", tk.END).strip()

    if not excel_path or not phone_col or not msg:
        messagebox.showerror("Error", "Please select Excel file, phone column, and enter a WhatsApp message.")
        return

    try:
        send_whatsapp_messages(excel_path, phone_col, msg, bool(wa_attach_pdf_var.get()), wa_pdf_default_var.get().strip(), wa_pdf_caption_var.get().strip())
        messagebox.showinfo("Done", "WhatsApp messages have been processed.")
        if remember_var.get():
            save_settings()
    except Exception as e:
        messagebox.showerror("Error", f"Error while sending WhatsApp:\n{e}")


# ================== GUI CALLBACKS – EMAIL TAB ==================
def browse_excel_email():
    # Use same shared variable (if user wants to change it from Email tab)
    path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        excel_path_var.set(path)


def start_email_sending():
    excel_path     = excel_path_var.get().strip()
    name_col       = name_col_var.get().strip()
    email_col      = email_col_var.get().strip()
    msg            = msg_email_text.get("1.0", tk.END).strip()
    sender_email   = sender_email_var.get().strip()
    sender_pass    = sender_password_var.get().strip()
    email_subject  = email_subject_var.get().strip() or DEFAULT_EMAIL_SUBJECT
    attach_certs   = bool(attach_cert_var.get())

    if not excel_path or not name_col or not email_col or not msg:
        messagebox.showerror("Error", "Please select Excel file, name column, email column, and enter an email message.")
        return
    if not sender_email or not sender_pass:
        messagebox.showerror("Error", "Please enter sender email and App Password.")
        return

    try:
        send_emails_with_certificates(
            excel_path, name_col, email_col, msg,
            sender_email, sender_pass, email_subject,
            attach_certs
        )
        messagebox.showinfo("Done", "Emails have been processed.")
        if remember_var.get():
            save_settings()
    except Exception as e:
        messagebox.showerror("Error", f"Error while sending emails:\n{e}")


# ================== GUI CALLBACKS – CERTIFICATES TAB ==================
def cert_browse_template():
    path = filedialog.askopenfilename(
        title="Select certificate template",
        filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")]
    )
    if path:
        cert_template_var.set(path)


def cert_browse_excel():
    path = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        cert_excel_var.set(path)


def cert_open_folder():
    os.makedirs(CERT_DIR, exist_ok=True)
    try:
        os.startfile(CERT_DIR)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open folder:\n{e}")


def cert_preview():
    template_path = cert_template_var.get().strip()
    excel_path    = cert_excel_var.get().strip()
    name_col      = cert_name_col_var.get().strip()

    if not template_path or not excel_path or not name_col:
        messagebox.showerror("Error", "Please select template, Excel, and name column.")
        return

    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file:\n{e}")
        return

    if name_col not in df.columns:
        messagebox.showerror("Error", f"Column '{name_col}' not found in Excel.")
        return

    name = ""
    for val in df[name_col]:
        s = str(val).strip()
        if s:
            name = s
            break
    if not name:
        messagebox.showerror("Error", "No valid names in this column.")
        return

    try:
        font_size = int(cert_font_size_var.get())
        y_ratio   = float(cert_y_ratio_var.get())
        x_offset  = int(cert_x_offset_var.get())
    except ValueError:
        messagebox.showerror("Error", "Font size, Y ratio and X offset must be numeric.")
        return

    is_arabic = bool(cert_arabic_var.get())

    try:
        font_path = get_default_font_path()
    except Exception as e:
        messagebox.showerror("Error", f"Font error:\n{e}")
        return

    try:
        img = draw_name_on_template(template_path, name, font_path,
                                    font_size, y_ratio, x_offset, is_arabic)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to draw preview:\n{e}")
        return

    os.makedirs(CERT_DIR, exist_ok=True)
    preview_path = os.path.join(CERT_DIR, "preview_certificate.png")
    img.save(preview_path, "PNG")
    try:
        os.startfile(preview_path)
        log(f"[PREVIEW] {preview_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open preview image:\n{e}")


def cert_generate_all():
    template_path = cert_template_var.get().strip()
    excel_path    = cert_excel_var.get().strip()
    name_col      = cert_name_col_var.get().strip()

    if not template_path or not excel_path or not name_col:
        messagebox.showerror("Error", "Please select template, Excel, and name column.")
        return

    try:
        font_size = int(cert_font_size_var.get())
        y_ratio   = float(cert_y_ratio_var.get())
        x_offset  = int(cert_x_offset_var.get())
    except ValueError:
        messagebox.showerror("Error", "Font size, Y ratio and X offset must be numeric.")
        return

    is_arabic = bool(cert_arabic_var.get())

    try:
        font_path = get_default_font_path()
    except Exception as e:
        messagebox.showerror("Error", f"Font error:\n{e}")
        return

    try:
        generate_all_certificates(template_path, excel_path, name_col,
                                  font_path, font_size, y_ratio, x_offset, is_arabic)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate certificates:\n{e}")


# ================== BUILD GUI ==================
ensure_fonts_dir_and_download()

root = tk.Tk()
root.title(APP_TITLE)
root.configure(bg=BG_MAIN)
root.state("zoomed")
root.resizable(True, True)

# Set window icon (if available)
try:
    ico_path = resource_path("logo.ico")
    if os.path.exists(ico_path):
        root.iconbitmap(ico_path)
except Exception:
    pass

# Load logo image (png/jpg) and resize, then show left of title
logo_img = None
try:
    logo_path = resource_path("logo.png")
    if not os.path.exists(logo_path):
        logo_path = resource_path("logo.jpg")
    if os.path.exists(logo_path):
        img = Image.open(logo_path)
        img.thumbnail((90, 90))
        logo_img = ImageTk.PhotoImage(img)
except Exception as e:
    log(f"[LOGO] Failed to load logo: {e}")
    logo_img = None

header_frame = tk.Frame(root, bg=BG_MAIN)
header_frame.pack(pady=10, fill="x")

header_inner = tk.Frame(header_frame, bg=BG_MAIN)
header_inner.pack(anchor="w", padx=20)

if logo_img is not None:
    tk.Label(header_inner, image=logo_img, bg=BG_MAIN).grid(row=0, column=0, rowspan=2, padx=(0, 14), sticky="w")

tk.Label(
    header_inner,
    text=APP_TITLE,
    font=("Segoe UI", 22, "bold"),
    fg=ACCENT,
    bg=BG_MAIN
).grid(row=0, column=1, sticky="w")

tk.Label(
    header_inner,
    text=APP_SUBTITLE,
    font=("Segoe UI", 10),
    fg="#B0B0B0",
    bg=BG_MAIN
).grid(row=1, column=1, sticky="w")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

whatsapp_tab = tk.Frame(notebook, bg=BG_PANEL)
email_tab    = tk.Frame(notebook, bg=BG_PANEL)
cert_tab     = tk.Frame(notebook, bg=BG_PANEL)

notebook.add(whatsapp_tab, text="WhatsApp")
notebook.add(email_tab,    text="Email")
notebook.add(cert_tab,     text="Certificates")

# ===== Shared variables =====
excel_path_var     = tk.StringVar()
name_col_var       = tk.StringVar(value="name")
phone_col_var      = tk.StringVar(value="phone")
email_col_var      = tk.StringVar(value="email")
sender_email_var   = tk.StringVar()
sender_password_var= tk.StringVar()
email_subject_var  = tk.StringVar(value=DEFAULT_EMAIL_SUBJECT)
remember_var       = tk.IntVar(value=1)
attach_cert_var    = tk.IntVar(value=1)
wa_attach_pdf_var  = tk.IntVar(value=0)
wa_pdf_default_var = tk.StringVar()
wa_pdf_caption_var = tk.StringVar()

# ================== WHATSAPP TAB UI ==================
wa_panel = tk.Frame(whatsapp_tab, bg=BG_PANEL)
wa_panel.pack(fill="both", expand=True, padx=10, pady=10)

# Row 0: Excel
tk.Label(wa_panel, text="Excel file:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=0, column=0, sticky="w", padx=10, pady=8)
tk.Entry(wa_panel, textvariable=excel_path_var, width=50,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=0, column=1, padx=5, pady=8, sticky="we")
tk.Button(wa_panel, text="Browse", command=browse_excel_whatsapp,
          bg=BTN_BG, fg=BTN_FG, relief="flat", padx=10, pady=4).grid(
    row=0, column=2, padx=10, pady=8)

# Row 1: Name column
tk.Label(wa_panel, text="Name column:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=1, column=0, sticky="w", padx=10, pady=5)
tk.Entry(wa_panel, textvariable=name_col_var, width=15,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=1, column=1, padx=5, pady=5, sticky="w")

# Row 2: Phone column
tk.Label(wa_panel, text="Phone column:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=2, column=0, sticky="w", padx=10, pady=5)
tk.Entry(wa_panel, textvariable=phone_col_var, width=15,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=2, column=1, padx=5, pady=5, sticky="w")


# Row 3: Optional PDF (Default)
tk.Checkbutton(
    wa_panel,
    text="Attach Default PDF (optional)",
    variable=wa_attach_pdf_var,
    bg=BG_PANEL,
    fg="#CCCCCC",
    selectcolor=BG_PANEL,
    activebackground=BG_PANEL,
    activeforeground="#FFFFFF",
).grid(row=3, column=1, padx=5, pady=(6, 2), sticky="w")

tk.Label(wa_panel, text="Default PDF:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=4, column=0, sticky="w", padx=10, pady=5)

tk.Entry(wa_panel, textvariable=wa_pdf_default_var, width=50,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=4, column=1, padx=5, pady=5, sticky="we")

tk.Button(wa_panel, text="Browse", command=browse_pdf_whatsapp,
          bg=BTN_BG, fg=BTN_FG, relief="flat", padx=10, pady=4).grid(
    row=4, column=2, padx=10, pady=5)

tk.Label(wa_panel, text="PDF caption:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=5, column=0, sticky="w", padx=10, pady=5)

tk.Entry(wa_panel, textvariable=wa_pdf_caption_var, width=50,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=5, column=1, padx=5, pady=5, sticky="we")

# Row 3: WhatsApp message
tk.Label(wa_panel, text="WhatsApp message:", bg=BG_PANEL, fg=ACCENT,
         font=("Segoe UI", 10, "bold")).grid(
    row=6, column=0, sticky="nw", padx=10, pady=(10, 5))

msg_wa_text = scrolledtext.ScrolledText(
    wa_panel, width=60, height=10,
    bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
    borderwidth=1, relief="solid", font=("Segoe UI", 10)
)
msg_wa_text.grid(row=6, column=1, padx=5, pady=(10, 10), sticky="nsew")

default_wa = (
    "السلام عليكم يا بشمهندس 🌟\n"
    "شرفتونا بالتسجيل في السيشن.\n"
    "منتظرينكم في مقر Energix Academy في دمياط الجديدة.\n"
    "للتأكيد على الحجز برجاء إرسال اسمك وكلمة (تم).\n"
    "مع تمنياتنا بالتوفيق 🌿"
)
msg_wa_text.insert("1.0", default_wa)

wa_btn_frame = tk.Frame(wa_panel, bg=BG_PANEL)
wa_btn_frame.grid(row=6, column=2, padx=10, pady=(10, 10), sticky="n")

tk.Button(
    wa_btn_frame,
    text="Send via WhatsApp",
    command=start_whatsapp_sending,
    bg=BTN_BG,
    fg=BTN_FG,
    relief="flat",
    padx=10,
    pady=6,
    font=("Segoe UI", 10, "bold")
).grid(row=0, column=0, padx=5, pady=3)

tk.Button(
    wa_btn_frame,
    text="Save WA Template",
    command=lambda: save_template(msg_wa_text, TEMPLATES_WA_DIR, "WhatsApp"),
    bg="#555555",
    fg="#FFFFFF",
    relief="flat",
    padx=8,
    pady=4,
    font=("Segoe UI", 9)
).grid(row=1, column=0, padx=5, pady=2)

tk.Button(
    wa_btn_frame,
    text="Load WA Template",
    command=lambda: load_template(msg_wa_text, TEMPLATES_WA_DIR, "WhatsApp"),
    bg="#555555",
    fg="#FFFFFF",
    relief="flat",
    padx=8,
    pady=4,
    font=("Segoe UI", 9)
).grid(row=2, column=0, padx=5, pady=2)

wa_panel.grid_columnconfigure(1, weight=1)
wa_panel.grid_rowconfigure(6, weight=1)

# ================== EMAIL TAB UI ==================
email_panel = tk.Frame(email_tab, bg=BG_PANEL)
email_panel.pack(fill="both", expand=True, padx=10, pady=10)

# Row 0: Excel (shared)
tk.Label(email_panel, text="Excel file (same sheet):", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=0, column=0, sticky="w", padx=10, pady=8)
tk.Entry(email_panel, textvariable=excel_path_var, width=50,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=0, column=1, padx=5, pady=8, sticky="we")
tk.Button(email_panel, text="Browse", command=browse_excel_email,
          bg=BTN_BG, fg=BTN_FG, relief="flat", padx=10, pady=4).grid(
    row=0, column=2, padx=10, pady=8)

# Row 1: Name column (shared)
tk.Label(email_panel, text="Name column:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=1, column=0, sticky="w", padx=10, pady=5)
tk.Entry(email_panel, textvariable=name_col_var, width=15,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=1, column=1, padx=5, pady=5, sticky="w")

# Row 2: Email column
tk.Label(email_panel, text="Email column:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=2, column=0, sticky="w", padx=10, pady=5)
tk.Entry(email_panel, textvariable=email_col_var, width=15,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=2, column=1, padx=5, pady=5, sticky="w")

# Row 3–5: sender email / password / subject
tk.Label(email_panel, text="Sender email:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=3, column=0, sticky="w", padx=10, pady=5)
tk.Entry(email_panel, textvariable=sender_email_var, width=30,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=3, column=1, padx=5, pady=5, sticky="w")

tk.Label(email_panel, text="App Password:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=4, column=0, sticky="w", padx=10, pady=5)
tk.Entry(email_panel, textvariable=sender_password_var, width=30, show="*",
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=4, column=1, padx=5, pady=5, sticky="w")

tk.Label(email_panel, text="Email subject:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=5, column=0, sticky="w", padx=10, pady=5)
tk.Entry(email_panel, textvariable=email_subject_var, width=45,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=5, column=1, padx=5, pady=5, sticky="w")

# Row 6: remember + attach certs
tk.Checkbutton(
    email_panel,
    text="Remember sender email and templates on this device (password is never saved)",
    variable=remember_var,
    bg=BG_PANEL,
    fg="#CCCCCC",
    selectcolor=BG_PANEL,
    activebackground=BG_PANEL,
    activeforeground="#FFFFFF",
).grid(row=6, column=1, padx=5, pady=5, sticky="w")

tk.Checkbutton(
    email_panel,
    text="Attach certificates from 'certificates' folder",
    variable=attach_cert_var,
    bg=BG_PANEL,
    fg="#CCCCCC",
    selectcolor=BG_PANEL,
    activebackground=BG_PANEL,
    activeforeground="#FFFFFF",
).grid(row=6, column=2, padx=5, pady=5, sticky="w")

# Row 7: Email message
tk.Label(email_panel, text="Email message (use {name}):", bg=BG_PANEL, fg=ACCENT,
         font=("Segoe UI", 10, "bold")).grid(
    row=7, column=0, sticky="nw", padx=10, pady=(10, 5))

msg_email_text = scrolledtext.ScrolledText(
    email_panel, width=60, height=10,
    bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
    borderwidth=1, relief="solid", font=("Segoe UI", 10)
)
msg_email_text.grid(row=7, column=1, padx=5, pady=(10, 10), sticky="nsew")

default_email = (
    "Dear Eng. {name},\n\n"
    "Thank you for attending our session at Energix Academy.\n"
    "Please find your certificate attached to this email.\n\n"
    "Best regards,\n"
    "Energix Academy"
)
msg_email_text.insert("1.0", default_email)

email_btn_frame = tk.Frame(email_panel, bg=BG_PANEL)
email_btn_frame.grid(row=7, column=2, padx=10, pady=(10, 10), sticky="n")

tk.Button(
    email_btn_frame,
    text="Send Emails",
    command=start_email_sending,
    bg="#28A745",
    fg="#FFFFFF",
    relief="flat",
    padx=10,
    pady=6,
    font=("Segoe UI", 10, "bold")
).grid(row=0, column=0, padx=5, pady=3)

tk.Button(
    email_btn_frame,
    text="Save Email Template",
    command=lambda: save_template(msg_email_text, TEMPLATES_EM_DIR, "Email"),
    bg="#555555",
    fg="#FFFFFF",
    relief="flat",
    padx=8,
    pady=4,
    font=("Segoe UI", 9)
).grid(row=1, column=0, padx=5, pady=2)

tk.Button(
    email_btn_frame,
    text="Load Email Template",
    command=lambda: load_template(msg_email_text, TEMPLATES_EM_DIR, "Email"),
    bg="#555555",
    fg="#FFFFFF",
    relief="flat",
    padx=8,
    pady=4,
    font=("Segoe UI", 9)
).grid(row=2, column=0, padx=5, pady=2)

email_panel.grid_columnconfigure(1, weight=1)
email_panel.grid_rowconfigure(7, weight=1)

# ================== CERTIFICATES TAB UI ==================
cert_template_var   = tk.StringVar()
cert_excel_var      = tk.StringVar()
cert_name_col_var   = tk.StringVar(value="name")
cert_font_size_var  = tk.StringVar(value="50")
cert_y_ratio_var    = tk.StringVar(value="0.40")
cert_x_offset_var   = tk.StringVar(value="0")
cert_arabic_var     = tk.IntVar(value=1)

cert_panel = tk.Frame(cert_tab, bg=BG_PANEL)
cert_panel.pack(fill="both", expand=True, padx=10, pady=10)

tk.Label(cert_panel, text="Certificate template:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=0, column=0, sticky="w", padx=10, pady=8)
tk.Entry(cert_panel, textvariable=cert_template_var, width=50,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=0, column=1, padx=5, pady=8, sticky="we")
tk.Button(cert_panel, text="Browse", command=cert_browse_template,
          bg=BTN_BG, fg=BTN_FG, relief="flat", padx=10, pady=4).grid(
    row=0, column=2, padx=10, pady=8)

tk.Label(cert_panel, text="Excel file:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=1, column=0, sticky="w", padx=10, pady=8)
tk.Entry(cert_panel, textvariable=cert_excel_var, width=50,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=1, column=1, padx=5, pady=8, sticky="we")
tk.Button(cert_panel, text="Browse", command=cert_browse_excel,
          bg=BTN_BG, fg=BTN_FG, relief="flat", padx=10, pady=4).grid(
    row=1, column=2, padx=10, pady=8)

tk.Label(cert_panel, text="Name column:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=2, column=0, sticky="w", padx=10, pady=5)
tk.Entry(cert_panel, textvariable=cert_name_col_var, width=15,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=2, column=1, padx=5, pady=5, sticky="w")

tk.Label(cert_panel, text="Font size:", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=3, column=0, sticky="w", padx=10, pady=5)
tk.Entry(cert_panel, textvariable=cert_font_size_var, width=10,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=3, column=1, padx=5, pady=5, sticky="w")

tk.Label(cert_panel, text="Name Y position (0–1):", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=4, column=0, sticky="w", padx=10, pady=5)
tk.Entry(cert_panel, textvariable=cert_y_ratio_var, width=10,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=4, column=1, padx=5, pady=5, sticky="w")

tk.Label(cert_panel, text="X offset (pixels):", bg=BG_PANEL, fg=FG_TEXT).grid(
    row=5, column=0, sticky="w", padx=10, pady=5)
tk.Entry(cert_panel, textvariable=cert_x_offset_var, width=10,
         bg=ENTRY_BG, fg=ENTRY_FG, insertbackground=FG_TEXT,
         borderwidth=1, relief="solid").grid(
    row=5, column=1, padx=5, pady=5, sticky="w")

tk.Checkbutton(
    cert_panel,
    text="Names are Arabic (RTL)",
    variable=cert_arabic_var,
    bg=BG_PANEL,
    fg="#CCCCCC",
    selectcolor=BG_PANEL,
    activebackground=BG_PANEL,
    activeforeground="#FFFFFF"
).grid(row=6, column=1, padx=5, pady=5, sticky="w")

cert_btn_frame = tk.Frame(cert_panel, bg=BG_PANEL)
cert_btn_frame.grid(row=7, column=1, pady=15, sticky="e")

tk.Button(
    cert_btn_frame, text="Preview first certificate",
    command=cert_preview, bg=BTN_BG, fg=BTN_FG,
    relief="flat", padx=10, pady=6, font=("Segoe UI", 10, "bold")
).grid(row=0, column=0, padx=5)

tk.Button(
    cert_btn_frame, text="Generate all certificates",
    command=cert_generate_all, bg="#28A745", fg="#FFFFFF",
    relief="flat", padx=10, pady=6, font=("Segoe UI", 10, "bold")
).grid(row=0, column=1, padx=5)

tk.Button(
    cert_panel, text="Open certificates folder",
    command=cert_open_folder, bg="#6C757D", fg="#FFFFFF",
    relief="flat", padx=10, pady=6
).grid(row=7, column=2, padx=10, pady=15)

cert_panel.grid_columnconfigure(1, weight=1)

# ===== Load settings after widgets are ready =====
load_settings()

root.mainloop()