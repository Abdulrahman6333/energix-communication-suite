# Energix Suite v4 - CLEAN (WhatsApp bulk sender + Optional PDF)
# UI: Tkinter. Automation: Selenium (WhatsApp Web).
# NOTE: Keep this file indented with SPACES (4) only.

import os
import re
import time
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# Pillow for logo (optional but recommended)
try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

# Selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================
# App constants / theme
# =========================
APP_TITLE = "Energix Suite v4"
APP_SUBTITLE = "Enhanced WhatsApp Automation Module"

BG_MAIN = "#0f131a"
BG_PANEL = "#151c24"
FG_MAIN = "#e6e6e6"
MUTED = "#b0b0b0"
ACCENT = "#00a3ff"
BTN_BG = "#0b75c9"
BTN_FG = "#ffffff"

DEFAULT_DELAY_SEC = 5


# =========================
# Helpers
# =========================
def resource_path(rel_path: str) -> str:
    """
    Works in dev and in PyInstaller onefile.
    """
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, rel_path)


def normalize_egypt_number(raw: str) -> str:
    """
    Normalize Egyptian numbers:
    - Keep digits only.
    - If starts with 0 and length 11 (e.g., 010...), convert to 20XXXXXXXXXX
    - If starts with 1 and length 10 (e.g., 10...), convert to 2010...
    - If already starts with 20 and length 12, keep.
    """
    if raw is None:
        return ""
    s = re.sub(r"\D", "", str(raw))
    if not s:
        return ""
    if s.startswith("20") and len(s) in (12, 13):  # sometimes extra leading 0 sneaks in
        # trim possible extra digits
        if len(s) == 13 and s.startswith("200"):
            s = "20" + s[3:]
        return s
    if s.startswith("0") and len(s) == 11:
        return "2" + s  # 20XXXXXXXXXXX
    if len(s) == 10 and s.startswith("1"):
        return "20" + s
    if len(s) == 11 and s.startswith("1"):
        return "20" + s
    # fallback: if it already looks like international without +
    if len(s) >= 11 and s.startswith("2"):
        return s
    return s


def safe_read_excel(path: str) -> pd.DataFrame:
    ext = Path(path).suffix.lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path)
    if ext in (".csv",):
        return pd.read_csv(path)
    raise ValueError("Unsupported file type. Please use .xlsx or .csv")


def ensure_spaces_only(text: str) -> str:
    # Not used at runtime; here as reminder.
    return text.replace("\t", "    ")


# =========================
# WhatsApp Automation
# =========================
def build_driver(chrome_profile_dir: str | None, headless: bool, logger):
    options = ChromeOptions()
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    if headless:
        # WhatsApp Web often breaks in headless; kept as optional.
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1400,900")

    if chrome_profile_dir:
        # Use a persistent profile to avoid QR each run
        options.add_argument(f'--user-data-dir={chrome_profile_dir}')
        # If multiple profiles, user can add profile-directory inside the path by setting it to that folder.

    driver = webdriver.Chrome(options=options)
    logger("[INFO] Chrome launched.")
    return driver


def wait_whatsapp_ready(driver, logger, timeout=120):
    """
    Wait until WhatsApp Web is ready (logged in).
    If QR is present, we wait up to timeout for login.
    """
    logger("[INFO] Opening WhatsApp Web...")
    driver.get("https://web.whatsapp.com/")
    t0 = time.time()

    while time.time() - t0 < timeout:
        time.sleep(1.0)

        # Case: logged in -> left chat list / search box appears
        try:
            # Search box is fairly stable
            driver.find_element(By.XPATH, "//div[@contenteditable='true' and @role='textbox' and @data-tab]")
            logger("[INFO] WhatsApp appears ready (textbox found).")
            return True
        except Exception:
            pass

        # Case: QR present -> canvas or QR container exists
        page = driver.page_source.lower()
        if "qr" in page and "keep me signed in" in page:
            logger("[INFO] Waiting for QR login...")
        elif "use whatsapp on your phone to scan the code" in page:
            logger("[INFO] Waiting for QR login...")
        else:
            logger("[INFO] Waiting for WhatsApp to load...")

    logger("[ERROR] WhatsApp not ready. QR not scanned / page not loaded within timeout.")
    return False


def maybe_click_continue_to_chat(driver, logger):
    """
    Sometimes the /send?phone= link opens an intermediate page.
    Try to click any 'Continue to Chat' / 'Use WhatsApp Web' buttons if present.
    """
    patterns = [
        "//a[contains(., 'Continue to Chat')]",
        "//a[contains(., 'Continue to chat')]",
        "//a[contains(., 'Use WhatsApp Web')]",
        "//a[contains(., 'use WhatsApp Web')]",
        "//button[contains(., 'Continue to Chat')]",
        "//button[contains(., 'Use WhatsApp Web')]",
    ]
    for xp in patterns:
        try:
            el = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, xp)))
            el.click()
            time.sleep(1.0)
            logger("[INFO] Clicked intermediate button.")
            return
        except Exception:
            continue


def chat_is_invalid(driver) -> bool:
    page = driver.page_source.lower()
    bad = [
        "phone number shared via url is invalid",
        "shared via url is invalid",
        "invalid phone number",
        "رقم الهاتف",
        "غير صالح",
    ]
    return any(b in page for b in bad)


def find_input_box(driver, timeout=40):
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "footer")))

    candidates = [
        (By.CSS_SELECTOR, "footer div[contenteditable='true'][role='textbox']"),
        (By.XPATH, "//footer//div[@contenteditable='true' and @role='textbox']"),
        (By.XPATH, "//div[@contenteditable='true' and (@role='textbox' or @data-tab)]"),
    ]

    last_err = None
    for by, sel in candidates:
        try:
            el = wait.until(EC.element_to_be_clickable((by, sel)))
            return el
        except Exception as e:
            last_err = e
    raise last_err


def paste_multiline(driver, text: str):
    """
    Paste multi-line message reliably:
    - Try clipboard paste.
    - Fallback: send keys line by line using SHIFT+ENTER for new lines.
    """
    msg = "" if text is None else str(text)
    try:
        import pyperclip
        pyperclip.copy(msg)
        active = driver.switch_to.active_element
        active.send_keys(Keys.CONTROL, "v")
        return
    except Exception:
        pass

    active = driver.switch_to.active_element
    lines = msg.split("\n")
    for i, line in enumerate(lines):
        active.send_keys(line)
        if i != len(lines) - 1:
            active.send_keys(Keys.SHIFT, Keys.ENTER)


def click_send_if_present(driver):
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send' or @data-testid='send']/ancestor::button"))
        )
        btn.click()
        return True
    except Exception:
        return False


def attach_pdf_file(driver, file_path: str, caption: str, logger):
    """
    Attach PDF (Document) and send.
    """
    if not file_path:
        logger("[PDF] Empty path, skipped.")
        return False

    p = os.path.abspath(str(file_path).strip())
    if not os.path.exists(p):
        logger(f"[PDF] File not found: {p}")
        return False

    wait = WebDriverWait(driver, 35)

    # Open attach menu
    try:
        clip = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//span[@data-icon='clip' or @data-testid='attach-menu-plus']/ancestor::button | //div[@title='Attach']"
        )))
        clip.click()
        time.sleep(0.5)
    except Exception as e:
        logger(f"[PDF] Can't open attach menu: {e}")
        return False

    # Click Document explicitly
    doc_xpaths = [
        "//span[contains(@data-icon,'document')]/ancestor::*[self::button or self::div][1]",
        "//button[@aria-label='Document']",
        "//div[@role='button' and .//span[contains(@data-icon,'document')]]",
        "//span[contains(@data-testid,'attach-document')]/ancestor::*[self::button or self::div][1]",
    ]
    clicked = False
    for xp in doc_xpaths:
        try:
            btn = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
            btn.click()
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        logger("[PDF] Document button not found; continuing with file input.")

    # Choose file input
    try:
        file_input = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        file_input.send_keys(p)
    except Exception as e:
        logger(f"[PDF] Can't find file input: {e}")
        return False

    # Wait for upload preview
    time.sleep(1.3)

    # Optional caption
    if caption:
        try:
            cap_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true' and (@role='textbox' or @data-tab)]"))
            )
            cap_box.click()
            cap_box.send_keys(caption)
        except Exception:
            pass

    # Send (wait longer)
    try:
        send_btn = WebDriverWait(driver, 90).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@data-icon='send' or @data-testid='send']/ancestor::button"))
        )
        send_btn.click()
        time.sleep(1.2)
        return True
    except Exception as e:
        logger(f"[PDF] Can't click send: {e}")
        return False


def send_whatsapp_bulk(
    excel_path: str,
    phone_column: str,
    message_template: str,
    delay_sec: int,
    attach_pdf: bool,
    pdf_default_path: str,
    pdf_column: str,
    pdf_caption: str,
    chrome_profile_dir: str | None,
    headless: bool,
    logger
):
    if not excel_path:
        raise ValueError("Please select an Excel/CSV file.")
    df = safe_read_excel(excel_path)
    if phone_column not in df.columns:
        raise ValueError(f"Phone column '{phone_column}' not found in file columns: {list(df.columns)}")

    has_pdf_col = attach_pdf and pdf_column and (pdf_column in df.columns)

    driver = None
    try:
        driver = build_driver(chrome_profile_dir, headless, logger)
        ready = wait_whatsapp_ready(driver, logger, timeout=150)
        if not ready:
            raise RuntimeError("WhatsApp is not ready. Please scan QR and try again (or use Chrome profile).")

        for idx, row in df.iterrows():
            raw_number = row.get(phone_column, "")
            phone = normalize_egypt_number(raw_number)
            if not phone or not phone.isdigit():
                logger(f"[SKIP] Invalid number: {raw_number} -> {phone}")
                continue

            chat_url = f"https://web.whatsapp.com/send?phone={phone}"
            driver.get(chat_url)
            time.sleep(2.0)
            maybe_click_continue_to_chat(driver, logger)
            time.sleep(1.0)

            if chat_is_invalid(driver):
                logger(f"[SKIP] WhatsApp invalid phone: {phone}")
                continue

            # Input box
            try:
                input_box = find_input_box(driver, timeout=40)
            except Exception as e:
                logger(f"[×] Can't find input box for {phone}: {e}")
                continue

            # Clear and paste
            try:
                input_box.click()
                time.sleep(0.3)
                input_box.send_keys(Keys.CONTROL, "a")
                input_box.send_keys(Keys.DELETE)
                time.sleep(0.2)

                paste_multiline(driver, message_template)

                # Send via Enter, then fallback to send button
                try:
                    input_box.send_keys(Keys.ENTER)
                except Exception:
                    pass

                # If still not sent, click send
                click_send_if_present(driver)

                logger(f"[✓] Message sent (attempt) to: {phone}")
            except Exception as e:
                logger(f"[×] Failed message send to {phone}: {e}")
                continue

            time.sleep(1.5)

            # PDF attach
            if attach_pdf:
                row_pdf = ""
                if has_pdf_col:
                    try:
                        val = row.get(pdf_column, "")
                        if pd.notna(val):
                            row_pdf = str(val).strip()
                    except Exception:
                        row_pdf = ""
                chosen_pdf = row_pdf or (pdf_default_path.strip() if pdf_default_path else "")
                if chosen_pdf:
                    ok = attach_pdf_file(driver, chosen_pdf, pdf_caption, logger)
                    if ok:
                        logger(f"[✓] PDF sent to: {phone}")
                    else:
                        logger(f"[×] PDF failed for: {phone}")
                else:
                    logger("[PDF] No PDF path provided (row/default).")

            # delay between contacts
            time.sleep(max(1, int(delay_sec)))

        logger("[DONE] All WhatsApp messages processed.")
    finally:
        try:
            if driver:
                driver.quit()
        except Exception:
            pass


# =========================
# UI
# =========================
class EnergixApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)

        # window icon
        try:
            ico_path = Path("logo.ico")
            if ico_path.exists():
                self.root.iconbitmap(str(ico_path))
        except Exception:
            pass

        self.root.configure(bg=BG_MAIN)
        self.root.geometry("1200x720")

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("TNotebook", background=BG_MAIN, borderwidth=0)
        style.configure("TNotebook.Tab", padding=(10, 6))
        style.map("TNotebook.Tab", background=[("selected", BG_PANEL)], foreground=[("selected", FG_MAIN)])

        # Header
        self._build_header()

        # Notebook
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=12, pady=12)

        self.whatsapp_tab = tk.Frame(self.nb, bg=BG_PANEL)
        self.email_tab = tk.Frame(self.nb, bg=BG_PANEL)
        self.cert_tab = tk.Frame(self.nb, bg=BG_PANEL)

        self.nb.add(self.whatsapp_tab, text="WhatsApp")
        self.nb.add(self.email_tab, text="Email")
        self.nb.add(self.cert_tab, text="Certificates")

        self._build_whatsapp_tab()
        self._build_placeholder(self.email_tab, "Email tab (coming soon)")
        self._build_placeholder(self.cert_tab, "Certificates tab (coming soon)")

    def _build_header(self):
        header = tk.Frame(self.root, bg=BG_MAIN)
        header.pack(fill="x", pady=(10, 6))

        inner = tk.Frame(header, bg=BG_MAIN)
        inner.pack(anchor="w", padx=20)

        # Logo small on left
        self.logo_img = None
        try:
            logo_path = "logo.png" if Path("logo.png").exists() else ("logo.jpg" if Path("logo.jpg").exists() else None)
            if logo_path and Image is not None:
                img = Image.open(logo_path)
                img.thumbnail((100, 100))
                self.logo_img = ImageTk.PhotoImage(img)
        except Exception:
            self.logo_img = None

        if self.logo_img is not None:
            tk.Label(inner, image=self.logo_img, bg=BG_MAIN).grid(row=0, column=0, rowspan=2, padx=(0, 14), sticky="w")

        tk.Label(inner, text=APP_TITLE, font=("Segoe UI", 22, "bold"), fg=ACCENT, bg=BG_MAIN).grid(row=0, column=1, sticky="w")
        tk.Label(inner, text=APP_SUBTITLE, font=("Segoe UI", 10), fg=MUTED, bg=BG_MAIN).grid(row=1, column=1, sticky="w")

    def _build_placeholder(self, parent, text):
        tk.Label(parent, text=text, bg=BG_PANEL, fg=MUTED, font=("Segoe UI", 12)).pack(pady=40)

    def _build_whatsapp_tab(self):
        pad = 14
        frm = tk.Frame(self.whatsapp_tab, bg=BG_PANEL)
        frm.pack(fill="both", expand=True, padx=pad, pady=pad)

        # Variables
        self.excel_path_var = tk.StringVar(value="")
        self.phone_col_var = tk.StringVar(value="phone")
        self.delay_var = tk.IntVar(value=DEFAULT_DELAY_SEC)

        self.attach_pdf_var = tk.BooleanVar(value=False)
        self.pdf_default_var = tk.StringVar(value="")
        self.pdf_col_var = tk.StringVar(value="pdf_path")
        self.pdf_caption_var = tk.StringVar(value="")

        self.chrome_profile_var = tk.StringVar(value="")
        self.headless_var = tk.BooleanVar(value=False)

        # Row 0: Excel file
        tk.Label(frm, text="Excel file:", bg=BG_PANEL, fg=FG_MAIN).grid(row=0, column=0, sticky="w", pady=(0, 6))
        excel_entry = tk.Entry(frm, textvariable=self.excel_path_var, bg="#0e1218", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat")
        excel_entry.grid(row=0, column=1, sticky="ew", pady=(0, 6), padx=(8, 8))
        tk.Button(frm, text="Browse", bg=BTN_BG, fg=BTN_FG, relief="flat", command=self._browse_excel).grid(row=0, column=2, sticky="e", pady=(0, 6))

        frm.grid_columnconfigure(1, weight=1)

        # Row 1: phone column + delay
        tk.Label(frm, text="Phone column:", bg=BG_PANEL, fg=FG_MAIN).grid(row=1, column=0, sticky="w", pady=(0, 6))
        tk.Entry(frm, textvariable=self.phone_col_var, width=20, bg="#0e1218", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat").grid(row=1, column=1, sticky="w", pady=(0, 6), padx=(8, 8))

        tk.Label(frm, text="Delay (sec):", bg=BG_PANEL, fg=FG_MAIN).grid(row=1, column=1, sticky="e", pady=(0, 6), padx=(8, 110))
        tk.Spinbox(frm, from_=1, to=60, textvariable=self.delay_var, width=6).grid(row=1, column=1, sticky="e", pady=(0, 6), padx=(0, 8))

        # Row 2: message box
        tk.Label(frm, text="WhatsApp message (multi-line):", bg=BG_PANEL, fg=FG_MAIN).grid(row=2, column=0, sticky="w", pady=(6, 6))
        self.msg_txt = tk.Text(frm, height=10, bg="#0e1218", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat", wrap="word")
        self.msg_txt.grid(row=3, column=0, columnspan=3, sticky="nsew")
        frm.grid_rowconfigure(3, weight=1)

        # PDF section
        pdf_box = tk.LabelFrame(frm, text="PDF attachment (optional)", bg=BG_PANEL, fg=FG_MAIN, bd=1, relief="groove")
        pdf_box.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(10, 8))
        pdf_box.grid_columnconfigure(1, weight=1)

        tk.Checkbutton(pdf_box, text="Attach PDF", variable=self.attach_pdf_var, bg=BG_PANEL, fg=FG_MAIN, activebackground=BG_PANEL, selectcolor=BG_PANEL).grid(row=0, column=0, sticky="w", padx=10, pady=(6, 4))

        tk.Label(pdf_box, text="Default PDF:", bg=BG_PANEL, fg=FG_MAIN).grid(row=1, column=0, sticky="w", padx=10, pady=4)
        tk.Entry(pdf_box, textvariable=self.pdf_default_var, bg="#0e1218", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat").grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=4)
        tk.Button(pdf_box, text="Browse PDF", bg=BTN_BG, fg=BTN_FG, relief="flat", command=self._browse_pdf).grid(row=1, column=2, padx=(0, 10), pady=4)

        tk.Label(pdf_box, text="PDF column:", bg=BG_PANEL, fg=FG_MAIN).grid(row=2, column=0, sticky="w", padx=10, pady=4)
        tk.Entry(pdf_box, textvariable=self.pdf_col_var, width=18, bg="#0e1218", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat").grid(row=2, column=1, sticky="w", padx=(0, 8), pady=4)

        tk.Label(pdf_box, text="Caption:", bg=BG_PANEL, fg=FG_MAIN).grid(row=3, column=0, sticky="w", padx=10, pady=(4, 8))
        tk.Entry(pdf_box, textvariable=self.pdf_caption_var, bg="#0e1218", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat").grid(row=3, column=1, sticky="ew", padx=(0, 8), pady=(4, 8))

        # Chrome profile
        prof_box = tk.LabelFrame(frm, text="Chrome profile (optional - avoids QR each time)", bg=BG_PANEL, fg=FG_MAIN, bd=1, relief="groove")
        prof_box.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        prof_box.grid_columnconfigure(1, weight=1)

        tk.Label(prof_box, text="User-data-dir:", bg=BG_PANEL, fg=FG_MAIN).grid(row=0, column=0, sticky="w", padx=10, pady=6)
        tk.Entry(prof_box, textvariable=self.chrome_profile_var, bg="#0e1218", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat").grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=6)
        tk.Button(prof_box, text="Browse", bg=BTN_BG, fg=BTN_FG, relief="flat", command=self._browse_profile).grid(row=0, column=2, padx=(0, 10), pady=6)

        tk.Checkbutton(prof_box, text="Headless (not recommended)", variable=self.headless_var, bg=BG_PANEL, fg=FG_MAIN, activebackground=BG_PANEL, selectcolor=BG_PANEL).grid(row=1, column=0, sticky="w", padx=10, pady=(0, 8))

        # Action + log
        action_row = tk.Frame(frm, bg=BG_PANEL)
        action_row.grid(row=6, column=0, columnspan=3, sticky="ew")
        action_row.grid_columnconfigure(0, weight=1)

        tk.Button(action_row, text="Send via WhatsApp", bg=BTN_BG, fg=BTN_FG, relief="flat", command=self._start_send).grid(row=0, column=1, sticky="e", pady=(0, 8))

        self.log_txt = tk.Text(frm, height=9, bg="#0b0f14", fg=FG_MAIN, insertbackground=FG_MAIN, relief="flat")
        self.log_txt.grid(row=7, column=0, columnspan=3, sticky="nsew")
        frm.grid_rowconfigure(7, weight=1)

        self._log("[READY] Select Excel file, write message, then Send.")

    def _log(self, msg: str):
        ts = time.strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        self.log_txt.insert("end", line)
        self.log_txt.see("end")
        self.root.update_idletasks()

    def _browse_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.xlsm *.csv")])
        if p:
            self.excel_path_var.set(p)

    def _browse_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self.pdf_default_var.set(p)

    def _browse_profile(self):
        p = filedialog.askdirectory()
        if p:
            self.chrome_profile_var.set(p)

    def _start_send(self):
        excel_path = self.excel_path_var.get().strip()
        phone_col = self.phone_col_var.get().strip()
        message_template = self.msg_txt.get("1.0", "end").rstrip("\n")
        delay_sec = int(self.delay_var.get())

        if not excel_path:
            messagebox.showerror("Missing", "Please choose an Excel/CSV file.")
            return
        if not message_template.strip():
            messagebox.showerror("Missing", "Please type a WhatsApp message.")
            return

        attach_pdf = bool(self.attach_pdf_var.get())
        pdf_default = self.pdf_default_var.get().strip()
        pdf_col = self.pdf_col_var.get().strip()
        pdf_caption = self.pdf_caption_var.get().strip()

        chrome_profile = self.chrome_profile_var.get().strip() or None
        headless = bool(self.headless_var.get())

        self._log("[START] Sending WhatsApp messages...")

        def worker():
            try:
                send_whatsapp_bulk(
                    excel_path=excel_path,
                    phone_column=phone_col,
                    message_template=message_template,
                    delay_sec=delay_sec,
                    attach_pdf=attach_pdf,
                    pdf_default_path=pdf_default,
                    pdf_column=pdf_col,
                    pdf_caption=pdf_caption,
                    chrome_profile_dir=chrome_profile,
                    headless=headless,
                    logger=self._log
                )
            except Exception as e:
                self._log("[ERROR] " + str(e))
                self._log(traceback.format_exc())

        threading.Thread(target=worker, daemon=True).start()


def main():
    root = tk.Tk()
    app = EnergixApp(root)
    root.mainloop()


if __name__ == "__main__":
    import sys
    main()