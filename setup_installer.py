# -*- coding: utf-8 -*-
"""
Illinois Tax Scraper – One-Click Setup
======================================
Double-click this EXE to set up everything needed for the Illinois Tax Scraper.
It will:
  1. Install a self-contained Python (no system Python required)
  2. Install xlwings, playwright, pywin32
  3. Download the Chromium browser (~150 MB, internet required once)
  4. Configure xlwings so Excel's VBA button can call Python
  5. Copy Tax il.xlsm and tax_scraper.py to your chosen folder

After setup, just open Tax il.xlsm and click "Fetch Illinois Tax Rates".
"""

import os
import sys
import glob
import shutil
import zipfile
import urllib.request
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
PYTHON_VERSION   = "3.11.9"
PYTHON_ZIP_URL   = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/python-{PYTHON_VERSION}-embed-amd64.zip"
GET_PIP_URL      = "https://bootstrap.pypa.io/get-pip.py"
DEFAULT_INSTALL  = r"C:\TaxIllinoisScraper"
APP_TITLE        = "Illinois Tax Scraper – Setup"

# Color palette (dark premium theme)
BG        = "#1a1a2e"
BG2       = "#16213e"
ACCENT    = "#e05d00"
ACCENT2   = "#f07830"
FG        = "#ffffff"
FG2       = "#b0b8c8"
SUCCESS   = "#4caf50"
ERR       = "#f44336"
CARD      = "#0f3460"
LOG_BG    = "#0d0d1a"
LOG_FG    = "#a0e0a0"

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def get_bundled(filename):
    """Return path to a file bundled with the EXE (or alongside this script)."""
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)


def download_file(url, dest_path, progress_cb=None):
    """Download a URL to dest_path, calling progress_cb(downloaded, total) if given."""
    req = urllib.request.urlopen(url, timeout=60)
    total = int(req.headers.get("Content-Length", 0))
    done  = 0
    chunk = 65536
    with open(dest_path, "wb") as f:
        while True:
            data = req.read(chunk)
            if not data:
                break
            f.write(data)
            done += len(data)
            if progress_cb:
                progress_cb(done, total)
    req.close()


def fix_embedded_python_pth(python_dir):
    """
    Enable site.py in the embedded Python distribution so pip and site-packages work.
    The .pth file has 'import site' commented out – we uncomment it.
    """
    pth_files = glob.glob(os.path.join(python_dir, "python3*.._pth")) or \
                glob.glob(os.path.join(python_dir, "python3*._pth"))
    if not pth_files:
        # Try generic glob
        pth_files = [f for f in os.listdir(python_dir) if f.endswith("._pth")]
        pth_files = [os.path.join(python_dir, f) for f in pth_files]
    for pth in pth_files:
        with open(pth, "r", encoding="utf-8") as fh:
            content = fh.read()
        if "#import site" in content:
            content = content.replace("#import site", "import site")
            with open(pth, "w", encoding="utf-8") as fh:
                fh.write(content)


def write_xlwings_conf(install_dir, python_exe):
    """Write xlwings.conf next to Tax il.xlsm so VBA RunPython finds the right Python."""
    conf_path = os.path.join(install_dir, "xlwings.conf")
    conf = (
        f"PYTHON_WIN,{python_exe}\n"
        f"PYTHONPATH,{install_dir}\n"
        f"USE_UDF_SERVER,False\n"
    )
    with open(conf_path, "w", encoding="utf-8") as fh:
        fh.write(conf)
    return conf_path


# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────

class InstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(bg=BG)
        self.root.geometry("700x600")
        self.root.resizable(False, False)
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

        self.install_dir = tk.StringVar(value=DEFAULT_INSTALL)
        self._cancelled = False
        self._done      = False

        self._build_ui()

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=ACCENT, height=6)
        hdr.pack(fill="x")

        title_frame = tk.Frame(self.root, bg=BG, pady=24)
        title_frame.pack(fill="x", padx=32)

        tk.Label(title_frame, text="🏛  Illinois Tax Scraper", font=("Segoe UI", 20, "bold"),
                 bg=BG, fg=FG).pack(anchor="w")
        tk.Label(title_frame, text="One-Click Setup — installs everything automatically.",
                 font=("Segoe UI", 11), bg=BG, fg=FG2).pack(anchor="w", pady=(4, 0))

        # ── Install path chooser ──────────────────────────────────────────────
        card = tk.Frame(self.root, bg=CARD, bd=0, padx=24, pady=20)
        card.pack(fill="x", padx=32)

        tk.Label(card, text="Install Folder", font=("Segoe UI", 10, "bold"),
                 bg=CARD, fg=FG2).pack(anchor="w")

        row = tk.Frame(card, bg=CARD)
        row.pack(fill="x", pady=(6, 0))

        self.path_entry = tk.Entry(row, textvariable=self.install_dir,
                                   font=("Consolas", 10), bg=BG2, fg=FG,
                                   insertbackground=FG, relief="flat",
                                   highlightthickness=1, highlightcolor=ACCENT,
                                   highlightbackground="#2a2a4a")
        self.path_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        browse_btn = tk.Button(row, text="Browse…", command=self._browse,
                               bg=CARD, fg=ACCENT2, font=("Segoe UI", 9),
                               relief="flat", cursor="hand2", padx=10, pady=4,
                               activebackground=BG2, activeforeground=ACCENT)
        browse_btn.pack(side="right")

        # ── Info bullets ──────────────────────────────────────────────────────
        info = tk.Frame(self.root, bg=BG, pady=16)
        info.pack(fill="x", padx=40)

        bullets = [
            ("🐍", "Installs Python 3.11 (self-contained, no system Python needed)"),
            ("📦", "Installs xlwings, playwright, and pywin32 packages"),
            ("🌐", "Downloads Chromium browser (~150 MB, internet required)"),
            ("📊", "Sets up Tax il.xlsm with VBA macros and Run button"),
            ("⚙️",  "Configures xlwings so Excel can call Python automatically"),
        ]
        for icon, text in bullets:
            r = tk.Frame(info, bg=BG)
            r.pack(anchor="w", pady=2)
            tk.Label(r, text=icon, font=("Segoe UI", 11), bg=BG, fg=FG, width=3).pack(side="left")
            tk.Label(r, text=text, font=("Segoe UI", 9), bg=BG, fg=FG2).pack(side="left")

        # ── Progress bar ──────────────────────────────────────────────────────
        prog_frame = tk.Frame(self.root, bg=BG, padx=32)
        prog_frame.pack(fill="x")

        self.step_label = tk.Label(prog_frame, text="Ready to install.",
                                   font=("Segoe UI", 9), bg=BG, fg=FG2, anchor="w")
        self.step_label.pack(fill="x", pady=(0, 6))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Tax.Horizontal.TProgressbar",
                         troughcolor=BG2, background=ACCENT,
                         thickness=10, borderwidth=0)
        self.progress = ttk.Progressbar(prog_frame, style="Tax.Horizontal.TProgressbar",
                                        length=636, mode="determinate")
        self.progress.pack(fill="x")

        # ── Log box ───────────────────────────────────────────────────────────
        log_frame = tk.Frame(self.root, bg=BG, padx=32, pady=10)
        log_frame.pack(fill="both", expand=True)

        self.log = tk.Text(log_frame, bg=LOG_BG, fg=LOG_FG,
                           font=("Consolas", 8), relief="flat",
                           state="disabled", height=8,
                           highlightthickness=0, wrap="word")
        scroll = tk.Scrollbar(log_frame, command=self.log.yview, bg=BG2, troughcolor=BG2)
        self.log.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_frame = tk.Frame(self.root, bg=BG, pady=12, padx=32)
        btn_frame.pack(fill="x")

        self.cancel_btn = tk.Button(btn_frame, text="Cancel", command=self._cancel,
                                    bg=BG2, fg=FG2, font=("Segoe UI", 10),
                                    relief="flat", padx=20, pady=8, cursor="hand2",
                                    activebackground=BG, activeforeground=FG)
        self.cancel_btn.pack(side="right", padx=(8, 0))

        self.install_btn = tk.Button(btn_frame, text="▶  Install Now",
                                     command=self._start_install,
                                     bg=ACCENT, fg=FG, font=("Segoe UI", 10, "bold"),
                                     relief="flat", padx=20, pady=8, cursor="hand2",
                                     activebackground=ACCENT2, activeforeground=FG)
        self.install_btn.pack(side="right")

        # Bottom accent bar
        tk.Frame(self.root, bg=ACCENT, height=3).pack(fill="x", side="bottom")

    def _browse(self):
        d = filedialog.askdirectory(title="Choose install folder", initialdir=self.install_dir.get())
        if d:
            self.install_dir.set(d)

    def _cancel(self):
        if self._done:
            self.root.destroy()
            return
        self._cancelled = True
        self.root.destroy()

    def _log(self, msg, color=None):
        self.log.configure(state="normal")
        tag = f"tag_{id(msg)}"
        if color:
            self.log.tag_configure(tag, foreground=color)
            self.log.insert("end", msg + "\n", tag)
        else:
            self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_step(self, text, pct=None):
        self.step_label.configure(text=text)
        if pct is not None:
            self.progress["value"] = pct
        self.root.update_idletasks()

    def _start_install(self):
        self.install_btn.configure(state="disabled", bg="#555")
        self.path_entry.configure(state="disabled")
        threading.Thread(target=self._run_install, daemon=True).start()

    def _run_install(self):
        install_dir  = self.install_dir.get().strip()
        internal_dir = os.path.join(install_dir, "_internal")
        python_dir   = os.path.join(internal_dir, "python")
        browsers_dir = os.path.join(internal_dir, "browsers")
        python_exe   = os.path.join(python_dir, "python.exe")

        try:
            # ── Step 1: Create directories ────────────────────────────────────
            self._set_step("📁  Creating install directories…", 2)
            os.makedirs(python_dir, exist_ok=True)
            os.makedirs(browsers_dir, exist_ok=True)
            self._log(f"Install folder: {install_dir}")

            # ── Step 2: Download Python embeddable ────────────────────────────
            self._set_step("🐍  Downloading Python 3.11 (embeddable)…", 5)
            self._log(f"Downloading Python from:\n  {PYTHON_ZIP_URL}")
            py_zip = os.path.join(internal_dir, "python.zip")

            def py_progress(done, total):
                if total:
                    pct = 5 + int((done / total) * 20)
                    mb  = done / 1_048_576
                    self._set_step(f"🐍  Downloading Python… {mb:.1f} MB", pct)

            download_file(PYTHON_ZIP_URL, py_zip, py_progress)
            self._log("Download complete. Extracting…")

            # ── Step 3: Extract Python ─────────────────────────────────────────
            self._set_step("📦  Extracting Python…", 26)
            with zipfile.ZipFile(py_zip, "r") as z:
                z.extractall(python_dir)
            os.remove(py_zip)
            fix_embedded_python_pth(python_dir)
            self._log(f"Python 3.11 extracted to: {python_dir}")

            # ── Step 4: Install pip ────────────────────────────────────────────
            self._set_step("⚙️   Installing pip…", 30)
            get_pip = os.path.join(internal_dir, "get-pip.py")
            self._log("Downloading get-pip.py…")
            download_file(GET_PIP_URL, get_pip)
            self._run_cmd([python_exe, get_pip, "--quiet"], "Installing pip")
            os.remove(get_pip)
            self._log("pip installed.")

            # ── Step 5: Install Python packages ───────────────────────────────
            self._set_step("📦  Installing xlwings, playwright, pywin32…", 38)
            self._log("Running: pip install xlwings playwright pywin32")
            self._run_cmd(
                [python_exe, "-m", "pip", "install",
                 "xlwings", "playwright", "pywin32", "--quiet"],
                "Installing packages"
            )
            self._log("Packages installed.")

            # ── Step 6: Fix pywin32 DLLs for embedded Python ──────────────────
            self._set_step("🔧  Configuring pywin32…", 50)
            site_pkgs = os.path.join(python_dir, "Lib", "site-packages")
            sys32_src = os.path.join(site_pkgs, "pywin32_system32")
            if os.path.isdir(sys32_src):
                for dll in os.listdir(sys32_src):
                    if dll.endswith(".dll"):
                        shutil.copy2(os.path.join(sys32_src, dll), python_dir)
                self._log("pywin32 DLLs copied to Python root.")
            # Run post-install script
            post_install = os.path.join(site_pkgs, "win32", "lib", "pywin32_postinstall.py")
            if not os.path.exists(post_install):
                # fallback search
                for root, dirs, files in os.walk(site_pkgs):
                    if "pywin32_postinstall.py" in files:
                        post_install = os.path.join(root, "pywin32_postinstall.py")
                        break
            if os.path.exists(post_install):
                self._run_cmd([python_exe, post_install, "-install"], "pywin32 post-install", check=False)
                self._log("pywin32 post-install complete.")

            # ── Step 7: Install Playwright Chromium ───────────────────────────
            self._set_step("🌐  Downloading Chromium browser (~150 MB)…", 55)
            self._log("This will take 1-3 minutes depending on your internet speed…")
            env = os.environ.copy()
            env["PLAYWRIGHT_BROWSERS_PATH"] = browsers_dir
            self._run_cmd(
                [python_exe, "-m", "playwright", "install", "chromium"],
                "Installing Chromium",
                env=env
            )
            self._log("Chromium installed.")

            # ── Step 8: Copy application files ────────────────────────────────
            self._set_step("📁  Copying application files…", 88)
            scraper_src = get_bundled("tax_scraper.py")
            scraper_dst = os.path.join(install_dir, "tax_scraper.py")
            shutil.copy2(scraper_src, scraper_dst)
            self._log(f"Copied: tax_scraper.py → {scraper_dst}")

            xlsm_src = get_bundled("Tax il.xlsm")
            xlsm_dst = os.path.join(install_dir, "Tax il.xlsm")
            shutil.copy2(xlsm_src, xlsm_dst)
            self._log(f"Copied: Tax il.xlsm → {xlsm_dst}")

            # ── Step 9: Write xlwings.conf ────────────────────────────────────
            self._set_step("⚙️   Configuring xlwings…", 94)
            conf_path = write_xlwings_conf(install_dir, python_exe)
            self._log(f"xlwings.conf written: {conf_path}")

            # ── Step 10: Done! ─────────────────────────────────────────────────
            self._set_step("✅  Setup complete!", 100)
            self._log("")
            self._log("═" * 60, ACCENT2)
            self._log("  Setup complete! ✅", SUCCESS)
            self._log(f"  Open this file to start:  Tax il.xlsm", SUCCESS)
            self._log(f"  Location: {xlsm_dst}", SUCCESS)
            self._log("  Click the orange 'Fetch Illinois Tax Rates' button.", SUCCESS)
            self._log("═" * 60, ACCENT2)

            self._done = True
            self.root.after(0, self._on_done, xlsm_dst)

        except Exception as exc:
            self._log(f"\n❌ ERROR: {exc}", ERR)
            self._set_step("❌  Setup failed. See log above.", None)
            self.root.after(0, lambda: messagebox.showerror(
                "Setup Failed",
                f"An error occurred during setup:\n\n{exc}\n\nPlease check your internet connection and try again."
            ))
            self.root.after(0, lambda: self.install_btn.configure(
                state="normal", bg=ACCENT, text="▶  Retry"))

    def _run_cmd(self, cmd, label, env=None, check=True):
        """Run a subprocess command, logging its output live."""
        self._log(f"  → {label}…")
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            env=env or os.environ.copy(),
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        )
        for line in proc.stdout:
            line = line.rstrip()
            if line:
                self._log(f"    {line}")
        proc.wait()
        if check and proc.returncode != 0:
            raise RuntimeError(f"Command failed (exit {proc.returncode}): {' '.join(cmd[:2])}")

    def _on_done(self, xlsm_path):
        self.cancel_btn.configure(text="Close")
        self.install_btn.configure(state="disabled", bg="#333", text="✅  Installed")
        result = messagebox.askyesno(
            "Setup Complete!",
            f"✅ Illinois Tax Scraper is ready!\n\n"
            f"Would you like to open Tax il.xlsm now?\n\n"
            f"File: {xlsm_path}"
        )
        if result:
            os.startfile(xlsm_path)
        self.root.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    app  = InstallerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
