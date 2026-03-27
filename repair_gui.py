import ctypes
import sys
import json
import os

def SetDPIAware():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

SetDPIAware()

import time
import threading
import win32gui
import win32ui
import win32con
from datetime import datetime
from PIL import Image
import tkinter as tk
from tkinter import font as tkfont
from pywinauto import Application, Desktop

# ============================================================
#  CONFIG
# ============================================================
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

def LoadConfig():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    config_path = os.path.join(base_dir, "config.json")

    if not os.path.exists(config_path):
        default = {
            "PHENOMENON_VALUE": "Appearance",
            "FAILURE_CODE":     "F173",
            "LOCATION":         "C801",
            "DUTY_CODE":        "Process",
            "REASON_CODE":      "SOLDERING--SOLDERING",
            "HANDLING":         "Touchup",
            "DUTY_DEPARTMENT":  "ME"
        }
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default, f, indent=4, ensure_ascii=False)
        return default

    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)

# ============================================================
#  COLOR PALETTE
# ============================================================
BG_DARK   = "#0e0f11"
BG_PANEL  = "#16181c"
BG_INPUT  = "#1e2128"
BORDER    = "#2a2d35"
AMBER     = "#f5a623"
AMBER_DIM = "#7a5212"
GREEN     = "#3ddc84"
RED_ERR   = "#ff4f4f"
TEXT_PRI  = "#e8eaf0"
TEXT_SEC  = "#6b7280"
TEXT_MONO = "#a8b0c0"
BLUE      = "#4ea8f5"


# ============================================================
#  DPI HELPER
# ============================================================
def GetDPI():
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return dpi
    except Exception:
        return 96


# ============================================================
#  GRAB WINDOW
# ============================================================
def GrabWindow(hwnd, rect=None):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win_w = right  - left
    win_h = bottom - top

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc  = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc_dc, win_w, win_h)
    save_dc.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)

    bmp_info = bmp.GetInfo()
    bmp_bits = bmp.GetBitmapBits(True)
    img = Image.frombuffer(
        "RGB",
        (bmp_info["bmWidth"], bmp_info["bmHeight"]),
        bmp_bits, "raw", "BGRX", 0, 1
    )

    win32gui.DeleteObject(bmp.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    if rect is not None:
        # ── relative crop จาก window ──────────────────────────
        # ไม่ใช้ absolute screen coord แต่ใช้ offset จาก window เอง
        crop_l = max(rect.left  - left, 0)
        crop_t = max(rect.top   - top,  0)
        crop_r = min(rect.right - left, win_w)
        crop_b = min(rect.bottom- top,  win_h)
        img = img.crop((crop_l, crop_t, crop_r, crop_b))
        print(f"[DEBUG] Crop rel=({crop_l},{crop_t},{crop_r},{crop_b}) "
              f"size={img.width}x{img.height}px")

    img.save("debug_grid.png")
    return img

# ============================================================
#  COLOR HELPER
# ============================================================
def is_red_bg(r, g, b):
    """
    Dynamic red detection
    """
    if r < 80:
        return False
    total = r + g + b
    if total < 100:
        return False

    # red channel ต้องครอง % สูงสุด
    r_ratio = r / total
    if r_ratio < 0.50:          # แดงต้องเป็น 50%+ ของสี
        return False
    if g / total > 0.30:        # เขียวต้องน้อย
        return False
    if b / total > 0.30:        # น้ำเงินต้องน้อย
        return False
    if r < g + 30:              # แดงต้องมากกว่าเขียวชัดเจน
        return False
    if r < b + 30:              # แดงต้องมากกว่าน้ำเงินชัดเจน
        return False

    return True


# ============================================================
#  WINDOW HELPERS
# ============================================================
def WaitForMainForm(timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        for w in Desktop(backend="win32").windows(title_re=r"^Repair-Rev"):
            if w.class_name() == "TfrmMain":
                return w
        time.sleep(0.5)
    raise RuntimeError(f"Timeout {timeout}s: TfrmMain did not appear")


def WaitForRepairWindow(timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        wins = Desktop(backend="win32").windows(title_re=r"^Repair Window")
        if wins:
            return wins[0]
        time.sleep(0.5)
    raise RuntimeError(f"Timeout {timeout}s: Repair Window did not appear")


# ============================================================
#  FIND CONTROLS
# ============================================================
def FindErrorCodeDBGrid(main_form):
    all_grids = []
    for child in main_form.descendants():
        try:
            if child.class_name() == "TDBGrid":
                all_grids.append(child)
        except Exception:
            pass

    if not all_grids:
        return None

    for child in main_form.descendants():
        try:
            if child.class_name() == "TBitBtn" and child.texts()[0] in ("New", "Remove"):
                btn_rect = child.rectangle()
                for grid in all_grids:
                    r = grid.rectangle()
                    if r.bottom <= btn_rect.top and abs(r.left - btn_rect.left) < 50:
                        return grid
        except Exception:
            pass

    leftmost = min(all_grids, key=lambda g: g.rectangle().left)
    return leftmost


def FindErrorCodeEdit(main_form):
    all_edits = []
    for child in main_form.descendants():
        try:
            if child.class_name() == "TDBEdit" and child.is_visible():
                all_edits.append(child)
        except Exception:
            pass

    for edit in all_edits:
        try:
            text = edit.texts()[0].strip()
            if not text:
                continue
            if text.startswith('F') or len(text) <= 12:
                if "height" in text.lower() or " " in text:
                    continue
                return edit
        except Exception:
            pass

    if len(all_edits) > 0:
        return all_edits[0]

    return None


# ============================================================
#  REPAIR WINDOW ACTIONS
# ============================================================
def SelectPhenomenon(repair_win_handle, value):
    try:
        app = Desktop(backend="win32")
        form = app.window(handle=repair_win_handle.handle)
        form.set_focus()
        time.sleep(0.3)

        all_combos = []
        for cb in form.descendants():
            try:
                if cb.class_name() == "TComboBox" and cb.is_visible():
                    r = cb.rectangle()
                    all_combos.append((cb, r))
            except Exception:
                pass

        if not all_combos:
            return False

        target_cb = min(all_combos, key=lambda x: (x[1].top, x[1].left))[0]
        r = target_cb.rectangle()
        target_cb.click_input(coords=(r.width() // 2, r.height() // 2))
        time.sleep(0.2)
        target_cb.set_focus()
        time.sleep(0.2)

        try:
            target_cb.select(value)
            return True
        except Exception:
            pass

        first_char = value[0] if value else ""
        if first_char:
            form.set_focus()
            time.sleep(0.2)
            form.type_keys(first_char)
            time.sleep(0.1)
            form.type_keys("{ENTER}")
            time.sleep(0.1)
            form.type_keys("{TAB}")
            return True

        return False

    except Exception as e:
        return False

def GetPanelEdits(form):
    """
    แยก TEdit ตาม Panel จริงๆ โดยใช้ left coordinate
    - Panel บน  (Failure Code / Failure Desc) → left ≈ 2614
    - Panel กลาง (LinkMonumber / Location / Part No.) → left ≈ 2610
    """
    form_rect = form.rectangle()
    form_mid  = form_rect.left + (form_rect.right - form_rect.left) // 2

    top_panel    = []  # Failure Code, Failure Desc
    middle_panel = []  # LinkMonumber, Location Code, Part No.

    for c in form.descendants():
        try:
            if c.class_name() == "TEdit" and c.is_visible():
                r = c.rectangle()
                if r.left >= form_mid:
                    continue  # ข้ามฝั่งขวา (Repairer, Memo)

                # แยกด้วย top — 2 panel อยู่คนละช่วง y
                # panel บน (Failure) top น้อย, panel กลาง (Location) top มากกว่า
                top_panel.append((c, r))
        except:
            pass

    top_panel.sort(key=lambda x: x[1].top)
    return [c for c, r in top_panel]


def FocusFailureCode(form):
    edits = GetPanelEdits(form)
    # index 0 = Failure Code (top น้อยสุด)
    if edits:
        edits[0].click_input()
        return edits[0]
    return None


def FocusLocationCode(form):
    edits = GetPanelEdits(form)
    # index 3 = Location Code (ข้าม Failure Code, Failure Desc, LinkMonumber)
    if len(edits) >= 4:
        edits[3].click_input()
        return edits[3]
    return None


def FillDutyCombos(form, duty_code, reason_code, handling, duty_department, log_fn):
    targets = []
    for c in form.descendants():
        try:
            if c.class_name() == "TComboBox" and c.is_visible():
                targets.append((c, c.rectangle()))
        except:
            pass

    targets.sort(key=lambda x: x[1].top)

    # index → field mapping จาก debug
    fields = [
        (2, duty_code,       "Duty Code"),
        (3, reason_code,     "Reason Code"),
        (4, handling,        "Handling"),
        (5, duty_department, "Duty Department"),
    ]

    for idx, value, label in fields:
        if idx >= len(targets):
            log_fn(f"  │  └ ✗ {label} index={idx} not found", RED_ERR)
            continue

        c = targets[idx][0]
        log_fn(f"  │ [{label}] = '{value}'", BLUE)
        try:
            c.select(value)
            log_fn(f"  │  └ ✓ Set via select()", GREEN)
        except Exception:
            try:
                c.click_input()
                time.sleep(0.1)
                c.type_keys(value[:2])
                time.sleep(0.1)
                c.type_keys("{ENTER}")
                log_fn(f"  │  └ ✓ Set via type_keys()", GREEN)
            except Exception as e:
                log_fn(f"  │  └ ✗ Failed: {e}", RED_ERR)

def ClickOK(form):
    try:
        form.child_window(title="OK", class_name="TBitBtn").click()
        time.sleep(0.3)
    except Exception as e:
        raise


# ============================================================
#  CORE
# ============================================================
def GetFirstRedErrorCode(main_form, phenomenon_value, sn,
                         failure_code, location,
                         duty_code, reason_code, handling, duty_department,
                         log_fn, status_fn):

    code          = None
    found_red_row = False

    # ── Step 1: หา Grid ─────────────────────────────────────
    status_fn("FINDING GRID", AMBER)
    log_fn("  ┌ Step 1/7 · Finding Error Code Grid...", BLUE)
    target_grid = FindErrorCodeDBGrid(main_form)
    if not target_grid:
        log_fn("  └ ✗ Error Code Grid not found", RED_ERR)
        return None, False
    log_fn(f"  └ ✓ Grid found at {target_grid.rectangle()}", GREEN)

    grid_rect = target_grid.rectangle()

    # ── Step 2: Capture ─────────────────────────────────────
    status_fn("CAPTURING", AMBER)
    log_fn("  ┌ Step 2/7 · Capturing window image...", BLUE)
    img = GrabWindow(main_form.handle, rect=grid_rect)
    pixels        = img.load()
    width, height = img.size
    log_fn(f"  └ ✓ Captured {width}x{height}px  (saved: debug_grid.png)", GREEN)

    # ── Step 3: Scan red row ─────────────────────────────────
    status_fn("SCANNING", AMBER)
    log_fn("  ┌ Step 3/7 · Scanning for red highlighted row...", BLUE)

    # sample RGB เพื่อ debug บนเครื่องอื่น
    log_fn("  │ Sampling pixel colors (first 60px):", BLUE)
    for y in range(0, min(height, 60), 6):
        sample_colors = []
        for x in range(0, width, max(width // 5, 1)):
            r, g, b = pixels[x, y]
            sample_colors.append(f"({r},{g},{b})")
        log_fn(f"  │  y={y:3d} → {' '.join(sample_colors)}", TEXT_SEC)

    first_red_y = None
    max_pct     = 0.0

    for y in range(height):
        red_count = sum(1 for x in range(width) if is_red_bg(*pixels[x, y]))
        pct       = red_count / width * 100
        if pct > max_pct:
            max_pct = pct
        if pct > 5.0 and first_red_y is None:
            first_red_y   = grid_rect.top + y
            found_red_row = True
            log_fn(f"  └ ✓ Red row at screen Y={first_red_y} "
                   f"({red_count}/{width}px = {pct:.1f}%)", GREEN)

    if not found_red_row:
        log_fn(f"  └ ℹ No red rows found "
               f"(max red% across all rows = {max_pct:.1f}%) → PASS", GREEN)
        return None, False

    # ── Step 4: Click red row ────────────────────────────────
    status_fn("SELECTING", AMBER)
    log_fn("  ┌ Step 4/7 · Clicking red row...", BLUE)
    main_rect = main_form.rectangle()
    click_x   = grid_rect.left + (grid_rect.right - grid_rect.left) // 2
    main_form.click_input(coords=(
        click_x     - main_rect.left,
        first_red_y - main_rect.top,
    ))
    time.sleep(0.5)
    log_fn(f"  └ ✓ Clicked at X={click_x}, Y={first_red_y}", GREEN)

    # ── Step 5: Read error code ──────────────────────────────
    status_fn("READING CODE", AMBER)
    log_fn("  ┌ Step 5/7 · Reading error code from TDBEdit...", BLUE)
    error_code_edit = FindErrorCodeEdit(main_form)
    if error_code_edit:
        raw_text = error_code_edit.texts()[0].strip()
        if raw_text:
            code = raw_text
            log_fn(f"  └ ✓ Error code: '{code}'", GREEN)
        else:
            log_fn("  └ ⚠ TDBEdit found but EMPTY", AMBER)
    else:
        log_fn("  └ ✗ TDBEdit not found", RED_ERR)

    # ── Step 6: Click Add ────────────────────────────────────
    status_fn("OPENING FORM", AMBER)
    log_fn("  ┌ Step 6/7 · Clicking Add button...", BLUE)
    try:
        add_btn = main_form.child_window(title="Add", class_name="TBitBtn")
        if not add_btn.exists():
            log_fn("  └ ⚠ Add button not found — skipping", AMBER)
            return code, found_red_row

        add_btn.click()
        log_fn("  │ ✓ Add clicked · waiting for Repair Window...", GREEN)

        repair_win = WaitForRepairWindow(timeout=5)
        log_fn("  └ ✓ Repair Window opened", GREEN)

        # ── Step 7: Fill Repair Window ───────────────────────
        log_fn("  ┌ Step 7/7 · Filling Repair Window...", BLUE)

        # 7a Phenomenon
        log_fn(f"  │ [7a] Phenomenon = '{phenomenon_value}'", BLUE)
        ok = SelectPhenomenon(repair_win, phenomenon_value)
        if ok:
            log_fn("  │  └ ✓ Phenomenon set", GREEN)
        else:
            log_fn("  │  └ ⚠ Phenomenon may not be set", AMBER)

        rapp  = Application(backend="win32").connect(handle=repair_win.handle)
        rform = rapp.window(handle=repair_win.handle)

        # 7b Failure Code
        log_fn(f"  │ [7b] Failure Code = '{failure_code}'", BLUE)
        edit = FocusFailureCode(rform)
        if edit:
            time.sleep(0.2)
            edit.type_keys("^a{BACKSPACE}")
            edit.type_keys(failure_code)
            edit.type_keys("{ENTER}")
            log_fn("  │  └ ✓ Failure Code set", GREEN)
        else:
            log_fn("  │  └ ✗ Failure Code field not found", RED_ERR)

        # 7c Location Code
        location_code = f"{sn}_{location}"
        log_fn(f"  │ [7c] Location Code = '{location_code}'", BLUE)
        loc_edit = FocusLocationCode(rform)
        if loc_edit:
            try:
                loc_edit.set_edit_text(location_code)
                log_fn("  │  └ ✓ Location Code set", GREEN)
            except Exception as e:
                log_fn(f"  │  └ ⚠ set_edit_text failed: {e} · trying clipboard", AMBER)
                import pyperclip
                pyperclip.copy(location_code)
                loc_edit.click_input()
                time.sleep(0.1)
                loc_edit.type_keys("^a^v")
                log_fn("  │  └ ✓ Location Code set via clipboard", GREEN)
        else:
            log_fn("  │  └ ✗ Location Code field not found", RED_ERR)

        # 7d Duty
        log_fn("  │ [7d] Filling Duty fields...", BLUE)
        FillDutyCombos(rform, duty_code, reason_code, handling, duty_department, log_fn)

        # 7d OK
        log_fn("  │ [7d] Clicking OK...", BLUE)
        ClickOK(rform)
        log_fn("  └ ✓ OK clicked · Repair Window closed", GREEN)

    except Exception as e:
        log_fn(f"  └ ✗ Add/Repair error: {e}", RED_ERR)

    return code, found_red_row

# ============================================================
#  MAIN PROCESS
# ============================================================
def RunRepairProcess(sn, log_fn, status_fn):
    try:
        cfg = LoadConfig()
        phenomenon_value = cfg.get("PHENOMENON_VALUE", "Appearance")
        failure_code     = cfg.get("FAILURE_CODE",     "F173")
        location         = cfg.get("LOCATION",         "C801")
        duty_code        = cfg.get("DUTY_CODE",        "Process")
        reason_code      = cfg.get("REASON_CODE",      "SOLDERING--SOLDERING")
        handling         = cfg.get("HANDLING",         "Touchup")
        duty_department  = cfg.get("DUTY_DEPARTMENT",  "ME")

        log_fn(f"· Config loaded:", BLUE)
        log_fn(f"  Phenomenon : {phenomenon_value}", TEXT_SEC)
        log_fn(f"  Failure    : {failure_code}", TEXT_SEC)
        log_fn(f"  Location   : {location}", TEXT_SEC)
        log_fn(f"  Duty Code  : {duty_code}", TEXT_SEC)
        log_fn(f"  Reason     : {reason_code}", TEXT_SEC)
        log_fn(f"  Handling   : {handling}", TEXT_SEC)
        log_fn(f"  Department : {duty_department}", TEXT_SEC)

        log_fn("━" * 42, AMBER)
        log_fn(f"  Serial Number : {sn}", TEXT_PRI)
        log_fn("━" * 42, AMBER)

        # ── หา Repair-Rev window ─────────────────────────────
        status_fn("CONNECTING", AMBER)
        log_fn("· Connecting to Repair-Rev window...", BLUE)
        repair_windows = Desktop(backend="win32").windows(
            title_re=r"^Repair-Rev",
            top_level_only=True,
            visible_only=True,
        )
        if not repair_windows:
            raise RuntimeError("Repair-Rev window not found")

        foreground_handle = ctypes.windll.user32.GetForegroundWindow()
        window = next(
            (w for w in repair_windows if w.handle == foreground_handle),
            repair_windows[0]
        )
        log_fn("✓ Connected to Repair-Rev", GREEN)

        app        = Application(backend="win32").connect(handle=window.handle)
        input_form = app.window(handle=window.handle)
        input_form.set_focus()

        # ── กรอก SN ──────────────────────────────────────────
        log_fn("· Filling serial number...", BLUE)
        sn_input = input_form.child_window(class_name="TEdit", found_index=0)
        sn_input.set_edit_text(sn)
        log_fn("✓ Serial number filled", GREEN)

        # ── Click Repair ─────────────────────────────────────
        log_fn("· Clicking Repair button...", BLUE)
        input_form.child_window(
            title="Repair", class_name="TBitBtn", found_index=0
        ).click()
        log_fn("✓ Repair clicked", GREEN)

        # ── รอ TfrmMain ──────────────────────────────────────
        status_fn("LOADING", AMBER)
        log_fn("· Waiting for main form...", BLUE)
        main_window = WaitForMainForm(timeout=10)
        repair_app  = Application(backend="win32").connect(handle=main_window.handle)
        repair_form = repair_app.window(handle=main_window.handle)
        repair_form.set_focus()
        time.sleep(1)
        log_fn("✓ Main form ready", GREEN)

        # ── Detect + Fill ─────────────────────────────────────
        status_fn("DETECTING", AMBER)
        log_fn("· Starting detection & repair flow...", BLUE)
        code, red_row_exists = GetFirstRedErrorCode(
            repair_form, phenomenon_value, sn,
            failure_code, location,
            duty_code, reason_code, handling, duty_department,
            log_fn, status_fn
        )

        # ── Result ───────────────────────────────────────────
        log_fn("━" * 42, AMBER)
        if red_row_exists:
            if code:
                log_fn(f"  RESULT · Error code found : {code}", RED_ERR)
                status_fn("ERROR FOUND", RED_ERR)
            else:
                log_fn("  RESULT · Red row detected but code UNREADABLE", RED_ERR)
                status_fn("READ ERROR", RED_ERR)
        else:
            log_fn("  RESULT · No errors found · PASS ✓", GREEN)
            status_fn("PASS", GREEN)
        log_fn("━" * 42, AMBER)

        return code

    except Exception as e:
        log_fn(f"✗ EXCEPTION : {e}", RED_ERR)
        status_fn("FAILED", RED_ERR)
        return None


# ============================================================
#  GUI
# ============================================================
class RepairGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Repair Automation")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(False, False)
        self.root.geometry("620x760")

        self._build_fonts()
        self._build_ui()
        self._set_ready()
        self.sn_entry.focus_set()

    def _build_fonts(self):
        self.f_title  = tkfont.Font(family="Consolas", size=13, weight="bold")
        self.f_label  = tkfont.Font(family="Consolas", size=9)
        self.f_entry  = tkfont.Font(family="Consolas", size=16, weight="bold")
        self.f_status = tkfont.Font(family="Consolas", size=22, weight="bold")
        self.f_log    = tkfont.Font(family="Consolas", size=9)
        self.f_btn    = tkfont.Font(family="Consolas", size=10, weight="bold")
        self.f_badge  = tkfont.Font(family="Consolas", size=8)

    def _build_ui(self):
        # ── header ──────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=BG_DARK)
        hdr.pack(fill="x", padx=20, pady=(20, 0))
        tk.Label(hdr, text="⬡  REPAIR AUTOMATION", font=self.f_title,
                 bg=BG_DARK, fg=AMBER).pack(side="left")
        self.clock_lbl = tk.Label(hdr, text="", font=self.f_badge,
                                  bg=BG_DARK, fg=TEXT_SEC)
        self.clock_lbl.pack(side="right", pady=4)
        self._tick_clock()

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x", padx=20, pady=(10, 0))

        # ── status badge ────────────────────────────────────
        sf = tk.Frame(self.root, bg=BG_PANEL,
                      highlightbackground=BORDER, highlightthickness=1)
        sf.pack(fill="x", padx=20, pady=(16, 0))
        inner = tk.Frame(sf, bg=BG_PANEL)
        inner.pack(pady=18)
        tk.Label(inner, text="STATUS", font=self.f_badge,
                 bg=BG_PANEL, fg=TEXT_SEC).pack()
        self.status_lbl = tk.Label(inner, text="READY", font=self.f_status,
                                   bg=BG_PANEL, fg=GREEN)
        self.status_lbl.pack()

        # ── SN input ─────────────────────────────────────────
        tk.Label(self.root, text="SERIAL NUMBER", font=self.f_badge,
                 bg=BG_DARK, fg=TEXT_SEC, anchor="w").pack(
                     fill="x", padx=20, pady=(16, 2))
        ef = tk.Frame(self.root, bg=AMBER, padx=2, pady=2)
        ef.pack(fill="x", padx=20)
        self.sn_var   = tk.StringVar()
        self.sn_entry = tk.Entry(
            ef, textvariable=self.sn_var,
            font=self.f_entry, bg=BG_INPUT, fg=TEXT_PRI,
            insertbackground=AMBER, relief="flat", bd=0)
        self.sn_entry.pack(fill="x", ipady=10, ipadx=12)
        self.sn_entry.bind("<Return>", self._on_scan)
        self.sn_entry.bind("<FocusOut>", lambda e: self.sn_entry.focus_set())
        tk.Label(self.root,
                 text="Scan barcode or type SN then press Enter",
                 font=self.f_badge, bg=BG_DARK, fg=TEXT_SEC).pack(
                     anchor="w", padx=20, pady=(4, 0))

        # ── result ───────────────────────────────────────────
        rf = tk.Frame(self.root, bg=BG_DARK)
        rf.pack(fill="x", padx=20, pady=(14, 0))
        tk.Label(rf, text="ERROR CODE", font=self.f_badge,
                 bg=BG_DARK, fg=TEXT_SEC).pack(anchor="w")
        self.result_lbl = tk.Label(rf, text="—", font=self.f_entry,
                                   bg=BG_DARK, fg=TEXT_SEC, anchor="w")
        self.result_lbl.pack(anchor="w")

        # ── buttons ──────────────────────────────────────────
        br = tk.Frame(self.root, bg=BG_DARK)
        br.pack(fill="x", padx=20, pady=(14, 0))
        self.reset_btn = tk.Button(
            br, text="↺  RESET", font=self.f_btn,
            bg=BG_PANEL, fg=AMBER,
            activebackground=AMBER_DIM, activeforeground=TEXT_PRI,
            relief="flat", bd=0, cursor="hand2",
            command=self._reset, padx=18, pady=8)
        self.reset_btn.pack(side="left")
        self.run_btn = tk.Button(
            br, text="▶  RUN", font=self.f_btn,
            bg=AMBER, fg=BG_DARK,
            activebackground=AMBER_DIM, activeforeground=TEXT_PRI,
            relief="flat", bd=0, cursor="hand2",
            command=lambda: self._on_scan(None), padx=18, pady=8)
        self.run_btn.pack(side="right")

        # ── log ──────────────────────────────────────────────
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x", padx=20, pady=(16, 0))
        lh = tk.Frame(self.root, bg=BG_DARK)
        lh.pack(fill="x", padx=20, pady=(8, 0))
        tk.Label(lh, text="PROCESS LOG", font=self.f_badge,
                 bg=BG_DARK, fg=TEXT_SEC).pack(side="left")
        tk.Button(lh, text="CLEAR", font=self.f_badge,
                  bg=BG_DARK, fg=TEXT_SEC, relief="flat", bd=0,
                  cursor="hand2", command=self._clear_log).pack(side="right")

        lf = tk.Frame(self.root, bg=BG_PANEL,
                      highlightbackground=BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True, padx=20, pady=(4, 20))
        self.log_text = tk.Text(
            lf, font=self.f_log, bg=BG_PANEL, fg=TEXT_MONO,
            insertbackground=AMBER, relief="flat", bd=0,
            state="disabled", wrap="word", padx=10, pady=8)
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll = tk.Scrollbar(
            lf, command=self.log_text.yview,
            bg=BG_PANEL, troughcolor=BG_PANEL,
            activebackground=BORDER, relief="flat", bd=0)
        scroll.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scroll.set)

        # tags
        self.log_text.tag_config("green",  foreground=GREEN)
        self.log_text.tag_config("red",    foreground=RED_ERR)
        self.log_text.tag_config("amber",  foreground=AMBER)
        self.log_text.tag_config("blue",   foreground=BLUE)
        self.log_text.tag_config("dim",    foreground=TEXT_SEC)
        self.log_text.tag_config("white",  foreground=TEXT_PRI)
        self.log_text.tag_config("normal", foreground=TEXT_MONO)

    def _tick_clock(self):
        self.clock_lbl.config(text=datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))
        self.root.after(1000, self._tick_clock)

    def _set_ready(self):
        self.running = False
        self.sn_entry.config(state="normal")
        self.run_btn.config(state="normal", bg=AMBER)
        self.reset_btn.config(state="normal")

    def _set_busy(self):
        self.running = True
        self.sn_entry.config(state="disabled")
        self.run_btn.config(state="disabled", bg=BORDER)
        self.reset_btn.config(state="disabled")

    def _log(self, message, color=None):
        def _write():
            ts  = datetime.now().strftime("%H:%M:%S")
            tag = {
                GREEN:    "green",
                RED_ERR:  "red",
                AMBER:    "amber",
                BLUE:     "blue",
                TEXT_PRI: "white",
            }.get(color, "normal")
            self.log_text.config(state="normal")
            self.log_text.insert("end", f"[{ts}] ", "dim")
            self.log_text.insert("end", f"{message}\n", tag)
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.root.after(0, _write)

    def _clear_log(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    def _set_status(self, text, color=GREEN):
        self.root.after(0, lambda: self.status_lbl.config(text=text, fg=color))

    def _set_result(self, text, color=TEXT_SEC):
        self.root.after(0, lambda: self.result_lbl.config(text=text, fg=color))

    def _on_scan(self, event):
        if self.running:
            return
        sn = self.sn_var.get().strip()
        if not sn:
            self._log("No serial number entered", RED_ERR)
            return

        self._set_busy()
        self._set_status("RUNNING...", AMBER)
        self._set_result("—", TEXT_SEC)

        def worker():
            code = RunRepairProcess(
                sn=sn,
                log_fn=self._log,
                status_fn=self._set_status,
            )
            if code:
                self.root.after(0, lambda: self._set_result(code, RED_ERR))
            else:
                self.root.after(0, lambda: self._set_result("No error found", GREEN))
            self.root.after(0, self._after_process)

        threading.Thread(target=worker, daemon=True).start()

    def _after_process(self):
        self.sn_var.set("")
        self._set_ready()
        self.sn_entry.focus_set()
        self._log("Ready for next scan", AMBER)

    def _reset(self):
        if self.running:
            return
        self.sn_var.set("")
        self._clear_log()
        self._set_status("READY", GREEN)
        self._set_result("—", TEXT_SEC)
        self.sn_entry.focus_set()
        self._log("System reset", AMBER)


# ============================================================
#  ENTRY POINT
# ============================================================
if __name__ == "__main__":
    root = tk.Tk()
    app  = RepairGUI(root)
    root.mainloop()