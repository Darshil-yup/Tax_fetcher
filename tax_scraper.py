# -*- coding: utf-8 -*-
import os
import sys

# ── Distribution auto-config ──────────────────────────────────────────────────
# When running from the installed distribution folder, point Playwright to the
# bundled Chromium so it doesn't need to be re-downloaded on each run.
_this_dir = os.path.dirname(os.path.abspath(__file__))
_browsers_dir = os.path.join(_this_dir, "_internal", "browsers")
if os.path.isdir(_browsers_dir):
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", _browsers_dir)
# ─────────────────────────────────────────────────────────────────────────────


import re
import time
import queue
import threading
import xlwings as xw
from playwright.sync_api import sync_playwright

class AddressNotFoundError(Exception):
    """Raised when the address is explicitly not found on the portal."""
    pass

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
MAX_WORKERS   = 10   # Increased to 10 parallel browsers for maximum throughput
MAX_RETRIES   = 2    # Auto-retry failed addresses this many times

# ---------------------------------------------------------------------------
# COLOR PALETTE  (Excel RGB ints: 0xBBGGRR for win32com)
# ---------------------------------------------------------------------------
SALES_HEADER_BG = 0xC45911
USE_HEADER_BG   = 0x3D7A16
SALES_SUB_BG    = 0xF4C19A
USE_SUB_BG      = 0xBCE6C4
SALES_TOTAL_BG  = 0xDD8C4E
USE_TOTAL_BG    = 0x6DB847
SALES_DATA_EVEN = 0xFAE8D8
SALES_DATA_ODD  = 0xFDF3EB
USE_DATA_EVEN   = 0xE6F5EC
USE_DATA_ODD    = 0xF2FAF5
HEADER_FONT     = 0xFFFFFF
SUBHEADER_FONT  = 0x1A1A1A
BUTTON_BG       = 0xC45911

# Fixed Excel output layout
SALES_TOTAL_COL = 6    # Column F
SALES_COLUMNS = {
    "State Tax": 7,                    # Column G
    "County Home Rule Tax": 8,         # Column H
    "County Public Safety": 9,         # Column I
    "County School Facility Tax": 10,  # Column J
    "Home Rule Tax": 11,               # Column K
    "Non-Home Rule Municipal Tax": 12, # Column L
    "RTA Tax": 13                      # Column M
}
GAP_COL = 14                           # Column N
USE_TOTAL_COL = 15                     # Column O
USE_COLUMNS = {
    "Local Use Tax": 16,               # Column P
    "State Use Tax": 17                # Column Q
}

# Locks for thread safety
print_lock   = threading.Lock()
results_lock = threading.Lock()

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def safe_print(*args, **kwargs):
    """Thread-safe logging wrapper."""
    with print_lock:
        print(*args, **kwargs)


def parse_pct(value_str):
    """'1.750%' -> 0.0175"""
    if not value_str:
        return 0.0
    try:
        return float(value_str.replace("%", "").strip()) / 100.0
    except ValueError:
        return 0.0


def fast_idle(page, timeout=12000):
    """Wait for the portal busy overlay to disappear, then a short settle."""
    try:
        page.wait_for_selector("div.FastBusyOverlay", state="hidden", timeout=timeout)
    except Exception:
        pass
    page.wait_for_timeout(75)


def navigate_to_form(page):
    """Go to the Tax Rate Finder > Search by Address form."""
    page.goto("https://mytax.illinois.gov/", wait_until="domcontentloaded")
    fast_idle(page)
    page.locator("a:has-text('Tax Rate Finder')").first.click()
    fast_idle(page)
    page.locator("a:has-text('Search by Address')").first.click()
    fast_idle(page)
    page.wait_for_selector("label:has-text('Street')", timeout=15000)


def strip_street_suffix(street):
    """Strip common suffixes (like St, Ave, Road, etc.) to optimize search matches."""
    street = street.strip()
    words = street.split()
    if len(words) > 1:
        last_word = words[-1].lower().rstrip('.')
        suffixes = {
            "st", "street", "rd", "road", "dr", "drive", "ave", "avenue",
            "ln", "lane", "blvd", "boulevard", "pl", "place", "ct", "court",
            "pkwy", "parkway", "way", "ter", "terrace"
        }
        if last_word in suffixes:
            return " ".join(words[:-1])
    return street


def scrape_breakdown(page):
    """
    Parse the Rate Breakdown page that is ALREADY loaded on `page`.
    Returns list of (label_str, rate_float).
    """
    splits = []
    try:
        page.wait_for_selector("h2:has-text('Rate Breakdown')", timeout=12000)
        fast_idle(page)
    except Exception:
        return splits
    rows = page.locator("tr.TDR")
    for i in range(rows.count()):
        row  = rows.nth(i)
        tds  = row.locator("td")
        if tds.count() == 2:
            label = tds.nth(0).inner_text().strip()
            val_s = tds.nth(1).inner_text().strip()
            if label and "%" in val_s:
                splits.append((label, parse_pct(val_s)))
    return splits


def apply_style(cell, bg, fg=None, bold=False, size=10, fmt=None):
    """One-shot cell styler for xlwings Range objects."""
    cell.api.Interior.Color = bg
    cell.font.name  = "Aptos Narrow"
    cell.font.size  = size
    cell.font.bold  = bold
    if fg is not None:
        cell.api.Font.Color = fg
    if fmt:
        cell.number_format = fmt
    cell.api.HorizontalAlignment = -4108   # center
    cell.api.VerticalAlignment   = -4108   # middle


# ---------------------------------------------------------------------------
# WORKER LOOP (Browser Reuse & Page Navigation)
# ---------------------------------------------------------------------------

def worker_loop(worker_id, q, output_q):
    """
    Worker thread that runs a single persistent browser session, reuses the search form
    via 'Back' button navigation, handles auto unit-selection, and retries with stripped street suffix.
    """
    with sync_playwright() as p:
        browser = None
        page = None

        def close_browser():
            nonlocal browser, page
            if page:
                try:
                    page.close()
                except Exception:
                    pass
                page = None
            if browser:
                try:
                    browser.close()
                except Exception:
                    pass
                browser = None

        def block_resources(route):
            # Abort image, font, and media requests to save network and rendering time
            if route.request.resource_type in ["image", "font", "media"]:
                route.abort()
            else:
                route.continue_()

        def init_browser():
            nonlocal browser, page
            close_browser()
            safe_print(f"  [Worker {worker_id}] Launching chromium...")
            browser = p.chromium.launch(
                headless=True,
                args=["--disable-gpu", "--no-sandbox",
                      "--disable-dev-shm-usage", "--disable-extensions",
                      "--disable-background-networking",
                      "--disable-background-timer-throttling"]
            )
            page = browser.new_page()
            page.route("**/*", block_resources)
            page.set_default_timeout(20000)
            safe_print(f"  [Worker {worker_id}] Navigating to Search by Address form...")
            navigate_to_form(page)

        try:
            init_browser()
        except Exception as e:
            safe_print(f"  [Worker {worker_id}] Initial browser launch failed: {e}")

        while True:
            try:
                item = q.get_nowait()
            except queue.Empty:
                break

            row_num = item["row"]
            street  = str(item["address"]).strip()
            city    = str(item["city"]).strip()
            zip_raw = item["zip"]
            zip_code = re.sub(r"\D", "", str(zip_raw).split(".")[0])[:5] if zip_raw else ""
            key     = (street.lower(), city.lower(), zip_code)

            safe_print(f"  [Worker {worker_id}] Processing: {street}, {city} {zip_code} (Row {row_num})")

            success = False
            address_not_found = False
            last_err = None

            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    # Lazily start/restore browser if closed by a previous attempt's failure
                    if not browser or not page:
                        init_browser()

                    # Verify page is on the street input form, otherwise navigate there
                    street_input = page.get_by_label("Street", exact=True).first
                    if street_input.count() == 0 or not street_input.is_visible():
                        safe_print(f"  [Worker {worker_id}] Form not active, reloading form...")
                        navigate_to_form(page)
                        street_input = page.get_by_label("Street", exact=True).first

                    # Clear and fill form
                    street_input.fill(street)
                    page.get_by_label("City",   exact=True).first.fill(city)
                    page.get_by_label("Zip",    exact=True).first.fill(zip_code)

                    # Validate address
                    page.locator("button:has-text('Validate Address')").first.click(force=True)
                    fast_idle(page)

                    try:
                        page.wait_for_selector(
                            "a:has-text('Select this address'), "
                            "button:has-text('Search'):not([disabled]), "
                            "button:has-text('OK')",
                            timeout=15000
                        )
                    except Exception:
                        pass

                    ok_btn      = page.locator("button:has-text('OK')").first
                    modal_links = page.locator("a:has-text('Select this address')")

                    # Case A: Address not found dialog - try suffix stripping retry
                    if ok_btn.count() > 0 and ok_btn.is_visible():
                        ok_btn.click(force=True)
                        fast_idle(page)
                        cancel = page.locator("button:has-text('Cancel')").first
                        if cancel.count() > 0 and cancel.is_visible():
                            cancel.click(force=True)
                            fast_idle(page)
                        
                        # Try stripped suffix retry
                        stripped = strip_street_suffix(street)
                        if stripped != street:
                            safe_print(f"  [Worker {worker_id}] No match for '{street}'. Retrying stripped: '{stripped}'")
                            page.get_by_label("Street", exact=True).first.fill(stripped)
                            page.locator("button:has-text('Validate Address')").first.click(force=True)
                            fast_idle(page)
                            
                            # Wait for outcomes again
                            page.wait_for_selector(
                                "a:has-text('Select this address'), "
                                "button:has-text('Search'):not([disabled]), "
                                "button:has-text('OK')",
                                timeout=18000
                            )
                            ok_btn      = page.locator("button:has-text('OK')").first
                            modal_links = page.locator("a:has-text('Select this address')")
                            if ok_btn.count() > 0 and ok_btn.is_visible():
                                ok_btn.click(force=True)
                                fast_idle(page)
                                cancel = page.locator("button:has-text('Cancel')").first
                                if cancel.count() > 0 and cancel.is_visible():
                                    cancel.click(force=True)
                                    fast_idle(page)
                                raise AddressNotFoundError("Address not found on portal (even after stripping suffix)")
                        else:
                            raise AddressNotFoundError("Address not found on portal")

                    # Case B: Address selection list / Multi-unit check
                    if modal_links.count() > 0:
                        selected = False
                        for m in range(modal_links.count()):
                            links = page.locator("a:has-text('Select this address')")
                            if m >= links.count():
                                break
                            links.nth(m).click()
                            fast_idle(page)

                            body = page.evaluate("() => document.body.innerText").lower()
                            needs_unit = any(kw in body for kw in
                                             ["select unit", "specify unit", "apartment",
                                              "unit number", "please select a unit", "unit range"])

                            if needs_unit:
                                # First check if table rows are present
                                grid_rows = page.locator("tr.TDRClickable")
                                if grid_rows.count() > 0:
                                    grid_rows.first.click(force=True)
                                    fast_idle(page)
                                    save = page.locator("button:has-text('Save')").first
                                    if save.count() > 0 and save.is_visible():
                                        save.click(force=True)
                                        fast_idle(page)
                                        selected = True
                                        break
                                else:
                                    # Try the combobox dropdown toggle
                                    toggle = page.locator(
                                        "a.ComboboxButton, button.ComboboxButton, "
                                        "[role='button']:has-text('Toggle Combobox')"
                                    ).first
                                    if toggle.count() > 0 and toggle.is_visible():
                                        toggle.click(force=True)
                                        fast_idle(page)
                                        page.wait_for_selector("li.ui-menu-item, [role='option']", timeout=8000)
                                        page.locator("li.ui-menu-item, [role='option']").first.click(force=True)
                                        fast_idle(page)
                                        save = page.locator("button:has-text('Save')").first
                                        if save.count() > 0 and save.is_visible():
                                            save.click(force=True)
                                            fast_idle(page)
                                            selected = True
                                            break
                                            
                                # If all failed, cancel modal and try other options
                                cancel = page.locator("button:has-text('Cancel')").first
                                if cancel.count() > 0 and cancel.is_visible():
                                    cancel.click(force=True)
                                    fast_idle(page)
                            else:
                                save = page.locator("button:has-text('Save')").first
                                if save.count() > 0 and save.is_visible():
                                    save.click(force=True)
                                    fast_idle(page)
                                selected = True
                                break

                        if not selected:
                            raise ValueError("All matching address choices required units, auto-selection failed")

                    # Case C: Validated directly (no dialog/modal)

                    # ── Search ───────────────────────────────────────────────────
                    page.wait_for_selector(
                        "button:has-text('Search'):not([disabled])", timeout=12000)
                    page.locator("button:has-text('Search')").first.click(force=True)
                    page.wait_for_selector("h2[id]:has-text('Tax Rates')", timeout=20000)
                    fast_idle(page)

                    # ── Sales Tax: Total & Splits ────────────────────────────────
                    sales_row = page.locator(
                        "tr:has-text('Sales Taxes'):has-text('General Merchandise')").first
                    sales_rate = sales_row.locator("a:has-text('%')").first.inner_text().strip()
                    sales_total = parse_pct(sales_rate)

                    # Go to Sales breakdown & scrape splits
                    sales_row.locator("a:has-text('%')").first.click(force=True)
                    sales_splits = scrape_breakdown(page)
                    
                    # Go back to Tax Rates summary page
                    page.locator("button:has-text('Back')").first.click(force=True)
                    page.wait_for_selector("h2:has-text('Tax Rates')", timeout=15000)
                    fast_idle(page)

                    # ── Use Tax: Total & Splits ──────────────────────────────────
                    use_row = page.locator(
                        "tr:has-text('Use Taxes'):has-text('General Merchandise')").first
                    use_rate = use_row.locator("a:has-text('%')").first.inner_text().strip()
                    use_total = parse_pct(use_rate)

                    # Go to Use breakdown & scrape splits
                    use_row.locator("a:has-text('%')").first.click(force=True)
                    use_splits = scrape_breakdown(page)
                    
                    # Go back to Tax Rates summary page
                    page.locator("button:has-text('Back')").first.click(force=True)
                    page.wait_for_selector("h2:has-text('Tax Rates')", timeout=15000)
                    fast_idle(page)

                    # Go back to the Search by Address form
                    page.locator("button:has-text('Back')").first.click(force=True)
                    page.wait_for_selector("label:has-text('Street')", timeout=15000)
                    fast_idle(page)

                    safe_print(f"  [OK] Worker {worker_id} - Row {row_num}: "
                               f"Sales={sales_rate} ({len(sales_splits)} splits) | "
                               f"Use={use_rate} ({len(use_splits)} splits)")

                    result_data = {
                        "sales_total":  sales_total, "sales_splits": sales_splits,
                        "use_total":    use_total,   "use_splits":   use_splits
                    }
                    output_q.put((row_num, key, result_data))
                    success = True
                    break

                except AddressNotFoundError as e:
                    safe_print(f"  [Address Not Found] Worker {worker_id} - Row {row_num}: {street}, {city} {zip_code}")
                    result_data = {
                        "sales_total": "Address not found", "sales_splits": [],
                        "use_total":   "Address not found", "use_splits":   []
                    }
                    output_q.put((row_num, key, result_data))
                    address_not_found = True
                    break

                except Exception as e:
                    last_err = e
                    safe_print(f"  [RETRY {attempt}] Worker {worker_id} - Row {row_num}: {e}")
                    close_browser()
                    if attempt < MAX_RETRIES:
                        time.sleep(3)

            if not success and not address_not_found:
                safe_print(f"  [ERR] Worker {worker_id} - Row {row_num} failed: {last_err}")
                result_data = {
                    "sales_total": 0.0, "sales_splits": [],
                    "use_total":   0.0, "use_splits":   []
                }
                output_q.put((row_num, key, result_data))

            q.task_done()

        close_browser()


# ---------------------------------------------------------------------------
# EXCEL WRITE FUNCTIONS (Executed on Main Thread)
# ---------------------------------------------------------------------------

def write_headers(ws, max_row):
    """Write beautiful, fixed headers to the active sheet and clear existing outputs."""
    with results_lock:
        ws.range(f"F3:Q{max_row + 2}").clear()
        try:
            ws.range(f"F3:Q{max_row + 2}").api.UnMerge()
        except Exception:
            pass

        # Freeze the first 4 rows (split below row 4)
        try:
            ws.range("A5").select()
            ws.book.app.api.ActiveWindow.FreezePanes = False
            ws.book.app.api.ActiveWindow.FreezePanes = True
        except Exception as e:
            safe_print(f"  [Freeze Panes Warning] Could not freeze panes: {e}")


        # ── Row 3: Section headers (merged) ──────────────────────────────────────
        hdr_s = ws.range((3, SALES_TOTAL_COL), (3, 13)) # F to M
        hdr_s.value = "Sales Tax"
        try:
            hdr_s.merge()
        except Exception:
            pass
        for c in range(SALES_TOTAL_COL, 14):
            apply_style(ws.range((3, c)), SALES_HEADER_BG, HEADER_FONT, bold=True, size=11)

        # Gap column N (14)
        apply_style(ws.range((3, GAP_COL)), 0xFFFFFF)

        hdr_u = ws.range((3, USE_TOTAL_COL), (3, 17)) # O to Q
        hdr_u.value = "Use Tax"
        try:
            hdr_u.merge()
        except Exception:
            pass
        for c in range(USE_TOTAL_COL, 18):
            apply_style(ws.range((3, c)), USE_HEADER_BG, HEADER_FONT, bold=True, size=11)

        # ── Row 4: Sub-headers ───────────────────────────────────────────────────
        c = ws.range((4, SALES_TOTAL_COL))
        c.value = "Total Sales Tax"
        apply_style(c, SALES_TOTAL_BG, HEADER_FONT, bold=True, size=10)

        for name, col in SALES_COLUMNS.items():
            c = ws.range((4, col))
            c.value = name
            apply_style(c, SALES_SUB_BG, SUBHEADER_FONT, bold=True, size=10)

        # Gap subheader
        apply_style(ws.range((4, GAP_COL)), 0xFFFFFF)

        c = ws.range((4, USE_TOTAL_COL))
        c.value = "Total Use Tax"
        apply_style(c, USE_TOTAL_BG, HEADER_FONT, bold=True, size=10)

        for name, col in USE_COLUMNS.items():
            c = ws.range((4, col))
            c.value = name
            apply_style(c, USE_SUB_BG, SUBHEADER_FONT, bold=True, size=10)

        # Autofit columns
        try:
            ws.range("F:Q").columns.autofit()
        except Exception:
            pass


def write_row(ws, r_num, data):
    """Write tax results for a single row to Excel with zebra-striping and formatting."""
    even = (r_num % 2 == 0)
    s_bg = SALES_DATA_EVEN if even else SALES_DATA_ODD
    u_bg = USE_DATA_EVEN   if even else USE_DATA_ODD

    with results_lock:
        try:
            # ── Sales Tax ─────────────────────────────────────────────────────
            # Sales total cell
            c_sales = ws.range((r_num, SALES_TOTAL_COL))
            sales_val = data["sales_total"]
            if isinstance(sales_val, str):
                c_sales.value = sales_val
                apply_style(c_sales, SALES_TOTAL_BG, HEADER_FONT, bold=True, size=10)
            else:
                c_sales.value = sales_val if (sales_val and sales_val > 0) else None
                apply_style(c_sales, SALES_TOTAL_BG, HEADER_FONT, bold=True, size=10, fmt="0.00%")

            # Sales splits
            sd = dict(data["sales_splits"])
            for name, col in SALES_COLUMNS.items():
                c = ws.range((r_num, col))
                if isinstance(sales_val, str):
                    c.value = None
                    apply_style(c, s_bg, size=10)
                else:
                    v = sd.get(name, 0.0)
                    c.value = v if (v and v > 0) else None
                    apply_style(c, s_bg, size=10, fmt="0.00%")

            # Gap column N (14)
            apply_style(ws.range((r_num, GAP_COL)), 0xFFFFFF, size=10)

            # ── Use Tax ───────────────────────────────────────────────────────
            # Use total cell
            c_use = ws.range((r_num, USE_TOTAL_COL))
            use_val = data["use_total"]
            if isinstance(use_val, str):
                c_use.value = use_val
                apply_style(c_use, USE_TOTAL_BG, HEADER_FONT, bold=True, size=10)
            else:
                c_use.value = use_val if (use_val and use_val > 0) else None
                apply_style(c_use, USE_TOTAL_BG, HEADER_FONT, bold=True, size=10, fmt="0.00%")

            # Use splits
            ud = dict(data["use_splits"])
            for name, col in USE_COLUMNS.items():
                c = ws.range((r_num, col))
                if isinstance(use_val, str):
                    c.value = None
                    apply_style(c, u_bg, size=10)
                else:
                    v = ud.get(name, 0.0)
                    c.value = v if (v and v > 0) else None
                    apply_style(c, u_bg, size=10, fmt="0.00%")
                
        except Exception as e:
            safe_print(f"  [Excel Write Error] Row {r_num}: {e}")


def setup_vba_and_button(wb):
    """Ensure the workbook is .xlsm, contains VBA macros, and has the Run button."""
    path = wb.fullname

    if path.lower().endswith(".xlsx"):
        new_path = path[:-5] + ".xlsm"
        safe_print(f"Converting to .xlsm: {new_path}")
        wb.api.SaveAs(new_path, 52)

    try:
        has_xw = any(c.Name == "xlwings" for c in wb.api.VBProject.VBComponents)
        if not has_xw:
            bas = os.path.join(os.path.dirname(xw.__file__), "xlwings.bas")
            if os.path.exists(bas):
                wb.api.VBProject.VBComponents.Import(bas)
    except Exception as e:
        safe_print(f"  xlwings.bas skipped: {e}")

    try:
        has_mod = any(c.Name == "TaxScraperModule"
                      for c in wb.api.VBProject.VBComponents)
        if not has_mod:
            mod = wb.api.VBProject.VBComponents.Add(1)
            mod.Name = "TaxScraperModule"
            mod.CodeModule.AddFromString(
                'Sub RunTaxScraper()\n'
                '    RunPython "import tax_scraper; tax_scraper.fetch_illinois_tax()"\n'
                'End Sub\n'
            )
    except Exception as e:
        safe_print(f"  VBA module skipped: {e}")

    try:
        ws = wb.sheets.active
        btn_name = "BtnRunScraper"
        for shape in ws.shapes:
            if shape.name == btn_name:
                shape.api.Delete()
                break
        r   = ws.range("A1:D2")
        btn = ws.shapes.api.AddShape(5, r.left, r.top, r.width, r.height)
        btn.Name     = btn_name
        btn.OnAction = "RunTaxScraper"
        try:
            btn.Adjustments.Item[1] = 0.12
        except Exception:
            pass
        btn.Fill.Solid()
        btn.Fill.ForeColor.RGB = BUTTON_BG
        btn.Line.Visible       = False
        btn.Shadow.Visible     = True
        btn.Shadow.OffsetX     = 2
        btn.Shadow.OffsetY     = 2
        btn.Shadow.Blur        = 4
        btn.Shadow.ForeColor.RGB = 0x808080
        tf = btn.TextFrame2
        tf.TextRange.Text = "Fetch Illinois Tax Rates"
        tf.TextRange.Font.Bold = True
        tf.TextRange.Font.Size = 12
        tf.TextRange.Font.Name = "Aptos Narrow"
        tf.TextRange.Font.Fill.ForeColor.RGB = 0xFFFFFF
        tf.HorizontalAnchor = 2
        tf.VerticalAnchor   = 3
        tf.TextRange.ParagraphFormat.Alignment = 2
        safe_print("Button created in A1:D2.")
    except Exception as e:
        safe_print(f"  Button error: {e}")

    wb.save()


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def fetch_illinois_tax():
    # ── Connect to workbook ──────────────────────────────────────────────────
    try:
        wb = xw.Book.caller()
        safe_print("Connected via xw.Book.caller()")
    except Exception:
        try:
            wb = xw.books.active
            safe_print("Connected to active workbook.")
        except Exception:
            try:
                app = xw.App(visible=True, add_book=False)
                fname = "Tax il.xlsm" if os.path.exists("Tax il.xlsm") else "Tax il.xlsx"
                wb = app.books.open(os.path.abspath(fname))
            except Exception as e:
                safe_print(f"Could not open workbook: {e}")
                return

    ws = wb.sheets.active
    safe_print(f"Sheet: {ws.name}")

    setup_vba_and_button(wb)

    # ── Read input rows (A-D, starting from row 5) in a single block ──────────
    try:
        last_row_a = int(ws.range("A1048576").end('up').row)
        last_row_b = int(ws.range("B1048576").end('up').row)
        last_row_d = int(ws.range("D1048576").end('up').row)
        last_row = max(5, last_row_a, last_row_b, last_row_d)
    except Exception:
        last_row = 100  # Safe fallback

    safe_print(f"Reading addresses from Row 5 to {last_row}...")
    values = ws.range(f"A5:D{last_row}").value
    if not values:
        values = []
    # Handle single-row response from xlwings
    if len(values) > 0 and not isinstance(values[0], list):
        values = [values]

    addresses = []
    for idx, val in enumerate(values):
        row_idx = 5 + idx
        if len(val) >= 4:
            a, c, s, z = val[0], val[1], val[2], val[3]
            # Add only non-empty rows
            if a or c or s or z:
                addresses.append({
                    "row": row_idx,
                    "address": a,
                    "city":    c,
                    "state":   s,
                    "zip":     z
                })

    if not addresses:
        safe_print("Sheet is empty - inserting demo address...")
        ws.range("A5").value = "3356 N Halsted St"
        ws.range("B5").value = "Chicago"
        ws.range("C5").value = "IL"
        ws.range("D5").value = 60657
        addresses.append({"row": 5, "address": "3356 N Halsted St",
                          "city": "Chicago", "state": "IL", "zip": 60657})

    # ── Deduplicate in-memory to build unique scrape list ──────────────────
    seen_keys = {}
    unique_to_scrape = []
    key_to_rows = {}

    for item in addresses:
        street  = str(item["address"]).strip()
        city    = str(item["city"]).strip()
        zip_raw = item["zip"]
        zip_code = re.sub(r"\D", "", str(zip_raw).split(".")[0])[:5] if zip_raw else ""
        key = (street.lower(), city.lower(), zip_code)
        
        if key not in key_to_rows:
            key_to_rows[key] = []
        key_to_rows[key].append(item["row"])

        if key not in seen_keys:
            seen_keys[key] = item
            unique_to_scrape.append(item)
        else:
            safe_print(f"  [DEDUP] Duplicate address row {item['row']} mapped to cached run.")

    total_unique = len(unique_to_scrape)
    workers      = min(MAX_WORKERS, total_unique)
    
    safe_print(f"\nProcessing {total_unique} unique address(es) (out of {len(addresses)} total rows)")
    safe_print(f"Spawning {workers} parallel browser workers...\n")

    # ── Write Beautiful Fixed Headers and Clear Old Outputs ─────────────────
    safe_print("Writing fixed sub-headers to Excel...")
    write_headers(ws, last_row)

    # ── Queue and run parallel scraper ────────────────────────────────────────
    q = queue.Queue()
    for item in unique_to_scrape:
        q.put(item)

    output_q = queue.Queue()

    # Spawn worker threads
    threads = []
    for w in range(workers):
        t = threading.Thread(target=worker_loop, args=(w + 1, q, output_q))
        t.start()
        threads.append(t)

    # ── Real-Time Dynamic Excel Writing (Executed in Main COM Thread) ────────
    while any(t.is_alive() for t in threads) or not output_q.empty():
        try:
            # Poll the output queue for completed results
            r_num, key, result = output_q.get(timeout=0.1)
            
            # Map result back to all matching rows (including duplicates) in real-time
            rows_to_write = key_to_rows.get(key, [r_num])
            for r in rows_to_write:
                safe_print(f"  [Writing Excel] Row {r}...")
                write_row(ws, r, result)
                
            output_q.task_done()
        except queue.Empty:
            continue

    # Ensure all threads are closed
    for t in threads:
        t.join()

    # Save and finalize
    ws.book.save()
    safe_print("\n[OK] Run completed successfully. Workbook saved.")


if __name__ == "__main__":
    fetch_illinois_tax()
