import ctypes
import sys
import json
import os
import csv
import zipfile
import xml.etree.ElementTree as ET

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
from tkinter import font as tkfont, filedialog, messagebox, ttk
from pywinauto import Application, Desktop

# ============================================================
#  CONFIG
# ============================================================
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")
LOG_DIR = os.path.join(os.path.dirname(__file__), "logs")
DEBUG_CONTROL_LOGS = False

def DefaultConfig():
    return {
        "MODE": "SCRAP",
        "DUTY_CODE": "Material",
        "REASON_CODE": "M01--Electrical",
        "HANDLING": "Scrap",
        "DUTY_DEPARTMENT": "VQA",
        "SCRAP_CODE": "0022 - Component Problem (Vendor)",
        "LOCATION_CODE": "DC20042-DC20043 FET FAIL",
        "COST_CENTER": "S2314SH1---Prod.-HP-Operation",
        "MEMO_TEMPLATE": "FET fail : {sn}_ON Semiconductor",
        "PRIVILEGE_EMP": "86047725",
        "PRIVILEGE_PASSWORD": "Phanya000"
    }

def LoadConfig():
    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    config_path = os.path.join(base_dir, "config.json")

    if not os.path.exists(config_path):
        default = DefaultConfig()
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default, f, indent=4, ensure_ascii=False)
        return default

    with open(config_path, "r", encoding="utf-8") as f:
        loaded = json.load(f)

    cfg = DefaultConfig()
    cfg.update(loaded)
    return cfg


def _xlsx_col_index(cell_ref):
    letters = ""
    for ch in cell_ref:
        if ch.isalpha():
            letters += ch.upper()
        else:
            break
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - ord("A") + 1)
    return value - 1


def _serials_from_rows(rows):
    if not rows:
        return []

    # Row 1 contains the work number. Serial numbers begin on row 2.
    data_rows = rows[1:]
    if not data_rows:
        return []

    header = [str(v).strip().lower().replace("_", " ") for v in data_rows[0]]
    wanted = ("serial number", "serialnumber", "sn", "serial")
    sn_col = None
    for idx, name in enumerate(header):
        compact = name.replace(" ", "")
        if name in wanted or compact in wanted:
            sn_col = idx
            break

    if sn_col is None:
        sn_col = 0
    else:
        data_rows = data_rows[1:]

    serials = []
    seen = set()
    for row in data_rows:
        if sn_col >= len(row):
            continue
        sn = str(row[sn_col]).strip()
        if not sn or sn.lower() in ("serial number", "serial", "sn"):
            continue
        if sn not in seen:
            serials.append(sn)
            seen.add(sn)
    return serials


def ReadSerialsFromExcel(path):
    ext = os.path.splitext(path)[1].lower()

    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            return _serials_from_rows(list(csv.reader(f)))

    if ext not in (".xlsx", ".xlsm"):
        raise ValueError("Please import a .xlsx, .xlsm, or .csv file")

    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    }

    with zipfile.ZipFile(path) as zf:
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("main:si", ns):
                shared.append("".join(t.text or "" for t in si.findall(".//main:t", ns)))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheet = workbook.find(".//main:sheet", ns)
        rel_id = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]

        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels:
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib["Target"]
                break
        if not target:
            raise ValueError("Could not find first worksheet")

        sheet_path = "xl/" + target.lstrip("/")
        sheet_xml = ET.fromstring(zf.read(sheet_path))

    rows = []
    for row in sheet_xml.findall(".//main:row", ns):
        values = []
        for cell in row.findall("main:c", ns):
            col = _xlsx_col_index(cell.attrib.get("r", "A1"))
            while len(values) <= col:
                values.append("")
            raw = cell.find("main:v", ns)
            if raw is None:
                inline = cell.find(".//main:t", ns)
                values[col] = (inline.text or "").strip() if inline is not None else ""
                continue
            value = raw.text or ""
            if cell.attrib.get("t") == "s":
                value = shared[int(value)] if value.isdigit() and int(value) < len(shared) else ""
            values[col] = value.strip()
        rows.append(values)

    return _serials_from_rows(rows)

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
    Dynamic red detection — ไม่ hardcode threshold
    ใช้ HSV-style ratio แทน absolute value
    รองรับ gamma/color profile ต่างกันทุกจอ
    """
    # แดงต้องสว่างพอ (ไม่ใช่ดำ)
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


def WaitForRepairWindowClosed(timeout=8):
    start = time.time()
    while time.time() - start < timeout:
        wins = Desktop(backend="win32").windows(
            title_re=r"^Repair Window",
            top_level_only=True,
            visible_only=True,
        )
        if not wins:
            return True
        time.sleep(0.2)
    return False


def WaitForWindowGone(hwnd, timeout=2):
    start = time.time()
    while time.time() - start < timeout:
        try:
            if not win32gui.IsWindow(hwnd):
                return True
            if not win32gui.IsWindowVisible(hwnd):
                return True
        except Exception:
            return True
        time.sleep(0.1)
    return False


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

def _class_matches(actual_class, expected_class):
    if actual_class == expected_class:
        return True
    lower = actual_class.lower()
    if expected_class == "TComboBox":
        return "combo" in lower
    if expected_class == "TEdit":
        return "edit" in lower
    return False


def GetVisibleControls(form, class_name):
    controls = []
    for c in form.descendants():
        try:
            if _class_matches(c.class_name(), class_name) and c.is_visible():
                controls.append((c, c.rectangle()))
        except Exception:
            pass
    controls.sort(key=lambda x: (x[1].top, x[1].left))
    return controls


def DumpScrapWindowControls(form, log_fn):
    log_fn("  + Scrap Window visible controls:", BLUE)
    for idx, c in enumerate(form.descendants()):
        try:
            if not c.is_visible():
                continue
            log_fn(f"  | #{idx:03d} {ControlDebugInfo(c)}", TEXT_SEC)
        except Exception:
            pass


def WaitForScrapControls(form, log_fn, timeout=5):
    start = time.time()
    while time.time() - start < timeout:
        combos = GetVisibleControls(form, "TComboBox")
        edits = GetVisibleControls(form, "TEdit")
        if len(combos) >= 6 and len(edits) >= 6:
            log_fn(f"  `- OK Controls ready: combos={len(combos)} edits={len(edits)}", GREEN)
            return True
        time.sleep(0.2)

    combos = GetVisibleControls(form, "TComboBox")
    edits = GetVisibleControls(form, "TEdit")
    log_fn(f"  `- WARN Controls not fully ready: combos={len(combos)} edits={len(edits)}", AMBER)
    return False


def _control_text(ctrl):
    try:
        texts = ctrl.texts()
        return texts[0].strip() if texts else ""
    except Exception:
        return ""


def ControlDebugInfo(ctrl):
    parts = []
    try:
        parts.append(f"class={ctrl.class_name()}")
    except Exception:
        parts.append("class=?")

    text = _control_text(ctrl)
    if text:
        parts.append(f"text='{text}'")

    try:
        parts.append(f"handle={ctrl.handle}")
    except Exception:
        pass

    try:
        control_id = ctrl.control_id()
        if control_id is not None:
            parts.append(f"id={control_id}")
    except Exception:
        pass

    try:
        r = ctrl.rectangle()
        parts.append(f"rect=({r.left},{r.top},{r.right},{r.bottom})")
    except Exception:
        pass

    return " ".join(parts)


def _norm_label(text):
    return text.lower().replace(":", "").replace(" ", "")


def FindControlByLabel(form, label_text, target_class):
    wanted = _norm_label(label_text)
    labels = []
    for c in form.descendants():
        try:
            if not c.is_visible():
                continue
            text = _control_text(c)
            if not text:
                continue
            normalized = _norm_label(text)
            if wanted == normalized or wanted in normalized:
                labels.append((c, c.rectangle()))
        except Exception:
            pass

    targets = GetVisibleControls(form, target_class)
    best = None
    best_score = None
    for _label, lr in labels:
        label_y = (lr.top + lr.bottom) // 2
        for target, tr in targets:
            target_y = (tr.top + tr.bottom) // 2
            if tr.left < lr.right - 5:
                continue
            dy = abs(target_y - label_y)
            if dy > 24:
                continue
            dx = max(tr.left - lr.right, 0)
            score = dy * 1000 + dx
            if best_score is None or score < best_score:
                best = target
                best_score = score

    return best


def FillComboByLabel(form, label_text, value, log_fn):
    combo = FindControlByLabel(form, label_text, "TComboBox")
    if not combo:
        log_fn(f"  |  `- FAIL {label_text} combo not found", RED_ERR)
        return False

    log_fn(f"  | [{label_text}] = '{value}'", BLUE)
    log_fn(f"  |  `- target label-match {ControlDebugInfo(combo)}", TEXT_SEC)
    try:
        combo.select(value)
        log_fn("  |  `- OK Set via select()", GREEN)
        return True
    except Exception:
        pass

    try:
        combo.click_input()
        time.sleep(0.1)
        import pyperclip
        pyperclip.copy(str(value))
        combo.type_keys("^a^v")
        time.sleep(0.1)
        combo.type_keys("{TAB}")
        log_fn("  |  `- OK Set via clipboard", GREEN)
        return True
    except Exception as e:
        log_fn(f"  |  `- FAIL {e}", RED_ERR)
        return False


def FillComboByLabelOrIndex(form, label_text, value, fallback_index, log_fn):
    if FillComboByLabel(form, label_text, value, log_fn):
        return True
    log_fn(f"  |  `- WARN using {label_text} fallback combo index {fallback_index}", AMBER)
    return FillComboByIndex(form, fallback_index, value, label_text, log_fn)


def FillEditByLabel(form, label_text, value, log_fn):
    edit = FindControlByLabel(form, label_text, "TEdit")
    if not edit:
        log_fn(f"  |  `- FAIL {label_text} edit not found", RED_ERR)
        return False

    log_fn(f"  | [{label_text}] = '{value}'", BLUE)
    log_fn(f"  |  `- target label-match {ControlDebugInfo(edit)}", TEXT_SEC)
    try:
        edit.set_edit_text(value)
        log_fn("  |  `- OK Set via set_edit_text()", GREEN)
        return True
    except Exception as e:
        log_fn(f"  |  `- WARN set_edit_text failed: {e}; trying clipboard", AMBER)
        try:
            import pyperclip
            pyperclip.copy(str(value))
            edit.click_input()
            time.sleep(0.1)
            edit.type_keys("^a^v")
            log_fn("  |  `- OK Set via clipboard", GREEN)
            return True
        except Exception as e2:
            log_fn(f"  |  `- FAIL {e2}", RED_ERR)
            return False


def _set_edit_text_native(edit, value):
    hwnd = getattr(edit, "handle", None)
    if not hwnd:
        return False

    text = str(value)
    win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, text)
    try:
        return win32gui.GetWindowText(hwnd) == text
    except Exception:
        return True


def FillEditControl(edit, value, label, log_fn):
    log_fn(f"  | [{label}] = '{value}'", BLUE)
    log_fn(f"  |  `- target {ControlDebugInfo(edit)}", TEXT_SEC)
    try:
        edit.set_edit_text(value)
        log_fn("  |  `- OK Set via set_edit_text()", GREEN)
        return True
    except Exception as e:
        log_fn(f"  |  `- WARN set_edit_text failed: {e}; trying native Win32", AMBER)
        try:
            if _set_edit_text_native(edit, value):
                log_fn("  |  `- OK Set via Win32 WM_SETTEXT", GREEN)
                return True
        except Exception as e2:
            log_fn(f"  |  `- WARN WM_SETTEXT failed: {e2}; trying keyboard", AMBER)

        try:
            edit.click_input()
            time.sleep(0.1)
            edit.type_keys("^a{BACKSPACE}")
            edit.type_keys(str(value), with_spaces=True)
            log_fn("  |  `- OK Set via keyboard", GREEN)
            return True
        except Exception as e2:
            log_fn(f"  |  `- FAIL {e2}", RED_ERR)
            return False


def FillLocationCodeSmart(form, value, log_fn):
    edit = FindControlByLabel(form, "Location Code", "TEdit")
    if edit:
        return FillEditControl(edit, value, "Location Code", log_fn)

    form_rect = form.rectangle()
    form_mid = form_rect.left + form_rect.width() // 2
    left_edits = [
        (edit, rect) for edit, rect in GetVisibleControls(form, "TEdit")
        if rect.left < form_mid
    ]
    # Scrap window left panel order: LinkMonumber, Location Code, Part No, ...
    if len(left_edits) > 1:
        log_fn("  |  `- WARN using Location Code fallback left edit index 1", AMBER)
        return FillEditControl(left_edits[1][0], value, "Location Code", log_fn)

    log_fn("  |  `- FAIL Location Code edit not found", RED_ERR)
    return False


def FillMemoSmart(form, value, log_fn):
    edit = FindControlByLabel(form, "Memo", "TEdit")
    if edit:
        return FillEditControl(edit, value, "Memo", log_fn)

    right_edits = GetRightStandaloneEdits(form)
    if right_edits:
        # Memo is below the two Repairer edits in the right panel.
        right_edits.sort(key=lambda item: (item[1].top, item[1].left))
        target = right_edits[2][0] if len(right_edits) > 2 else right_edits[-1][0]
        log_fn("  |  `- WARN using Memo fallback right_standalone_edit[2]", AMBER)
        return FillEditControl(target, value, "Memo", log_fn)

    log_fn("  |  `- FAIL Memo edit not found", RED_ERR)
    return False


def FillComboByIndex(form, index, value, label, log_fn):
    combos = GetVisibleControls(form, "TComboBox")
    if index >= len(combos):
        log_fn(f"  |  `- FAIL {label} combo index={index} not found", RED_ERR)
        return False

    combo = combos[index][0]
    log_fn(f"  | [{label}] = '{value}'", BLUE)
    log_fn(f"  |  `- target combo[{index}] {ControlDebugInfo(combo)}", TEXT_SEC)
    try:
        combo.select(value)
        log_fn("  |  `- OK Set via select()", GREEN)
        return True
    except Exception:
        try:
            combo.click_input()
            time.sleep(0.1)
            import pyperclip
            pyperclip.copy(str(value))
            combo.type_keys("^a^v")
            time.sleep(0.1)
            combo.type_keys("{ENTER}")
            time.sleep(0.1)
            combo.type_keys("{TAB}")
            log_fn("  |  `- OK Set via clipboard + enter", GREEN)
            return True
        except Exception as e:
            log_fn(f"  |  `- FAIL {e}", RED_ERR)
            return False


def FillLeftComboByIndex(form, index, value, label, log_fn):
    form_rect = form.rectangle()
    form_mid = form_rect.left + form_rect.width() // 2
    combos = [
        (combo, rect) for combo, rect in GetVisibleControls(form, "TComboBox")
        if rect.left < form_mid
    ]
    if index >= len(combos):
        log_fn(f"  |  `- FAIL {label} left combo index={index} not found", RED_ERR)
        return False

    combo = combos[index][0]
    log_fn(f"  | [{label}] = '{value}'", BLUE)
    log_fn(f"  |  `- target left_combo[{index}] {ControlDebugInfo(combo)}", TEXT_SEC)
    try:
        combo.select(value)
        log_fn("  |  `- OK Set via select()", GREEN)
        return True
    except Exception:
        try:
            combo.click_input()
            time.sleep(0.1)
            import pyperclip
            pyperclip.copy(str(value))
            combo.type_keys("^a^v")
            time.sleep(0.1)
            combo.type_keys("{ENTER}")
            time.sleep(0.1)
            combo.type_keys("{TAB}")
            log_fn("  |  `- OK Set via clipboard + enter", GREEN)
            return True
        except Exception as e:
            log_fn(f"  |  `- FAIL {e}", RED_ERR)
            return False


def FillLeftComboByPrefix(form, index, value, label, log_fn):
    form_rect = form.rectangle()
    form_mid = form_rect.left + form_rect.width() // 2
    combos = [
        (combo, rect) for combo, rect in GetVisibleControls(form, "TComboBox")
        if rect.left < form_mid
    ]
    if index >= len(combos):
        log_fn(f"  |  `- FAIL {label} left combo index={index} not found", RED_ERR)
        return False

    combo = combos[index][0]
    prefix = str(value).split("--", 1)[0].strip() or str(value)
    log_fn(f"  | [{label}] = '{value}'", BLUE)
    log_fn(f"  |  `- target left_combo[{index}] {ControlDebugInfo(combo)}", TEXT_SEC)

    try:
        combo.select(value)
        log_fn("  |  `- OK Set via select()", GREEN)
        return True
    except Exception:
        pass

    try:
        rect = combo.rectangle()
        combo.click_input(coords=(max(8, rect.width() // 4), rect.height() // 2))
        time.sleep(0.2)
        combo.type_keys("^a{BACKSPACE}")
        time.sleep(0.1)
        combo.type_keys(prefix, with_spaces=True, pause=0.05)
        time.sleep(0.2)
        combo.type_keys("{ENTER}")
        time.sleep(0.2)
        combo.type_keys("{TAB}")
        log_fn(f"  |  `- OK Set via prefix '{prefix}' + enter", GREEN)
        return True
    except Exception as e:
        log_fn(f"  |  `- WARN prefix failed: {e}; trying clipboard", AMBER)

    try:
        rect = combo.rectangle()
        combo.click_input(coords=(max(8, rect.width() // 4), rect.height() // 2))
        time.sleep(0.1)
        import pyperclip
        pyperclip.copy(prefix)
        combo.type_keys("^a^v")
        time.sleep(0.2)
        combo.type_keys("{ENTER}")
        time.sleep(0.2)
        combo.type_keys("{TAB}")
        log_fn(f"  |  `- OK Set via clipboard prefix '{prefix}' + enter", GREEN)
        return True
    except Exception as e:
        log_fn(f"  |  `- FAIL {e}", RED_ERR)
        return False


def FillLeftReasonCombo(form, index, value, log_fn):
    form_rect = form.rectangle()
    form_mid = form_rect.left + form_rect.width() // 2
    combos = [
        (combo, rect) for combo, rect in GetVisibleControls(form, "TComboBox")
        if rect.left < form_mid
    ]
    if index >= len(combos):
        log_fn(f"  |  `- FAIL Reason Code left combo index={index} not found", RED_ERR)
        return False

    combo, rect = combos[index]
    value = str(value)
    prefix = value.split("--", 1)[0].strip()
    log_fn(f"  | [Reason Code] = '{value}'", BLUE)
    log_fn(f"  |  `- target left_combo[{index}] {ControlDebugInfo(combo)}", TEXT_SEC)

    try:
        combo.select(value)
        log_fn("  |  `- OK Set via select()", GREEN)
        return True
    except Exception:
        pass

    try:
        combo.click_input(coords=(rect.width() - 10, rect.height() // 2))
        time.sleep(0.3)
        combo.type_keys("{HOME}")

        if len(prefix) >= 3 and prefix[0].upper() == "M" and prefix[1:].isdigit():
            down_count = max(int(prefix[1:]) - 1, 0)
            if down_count:
                combo.type_keys("{DOWN " + str(down_count) + "}")
            log_fn(f"  |  `- selecting dropdown row for {prefix}", BLUE)
        else:
            combo.type_keys(prefix, with_spaces=True, pause=0.05)
            log_fn(f"  |  `- selecting dropdown by prefix '{prefix}'", BLUE)

        time.sleep(0.1)
        combo.type_keys("{ENTER}")
        time.sleep(0.1)
        combo.type_keys("{TAB}")
        log_fn("  |  `- OK Set via dropdown selection", GREEN)
        return True
    except Exception as e:
        log_fn(f"  |  `- FAIL {e}", RED_ERR)
        return False


def GetStandaloneEdits(form):
    edits = []
    for c in form.descendants():
        try:
            if c.class_name() == "TEdit" and c.is_visible():
                edits.append((c, c.rectangle()))
        except Exception:
            pass
    edits.sort(key=lambda item: (item[1].top, item[1].left))
    return edits


def GetLeftStandaloneEdits(form):
    form_rect = form.rectangle()
    form_mid = form_rect.left + form_rect.width() // 2
    return [
        (edit, rect) for edit, rect in GetStandaloneEdits(form)
        if rect.left < form_mid
    ]


def GetRightStandaloneEdits(form):
    form_rect = form.rectangle()
    form_mid = form_rect.left + form_rect.width() // 2
    return [
        (edit, rect) for edit, rect in GetStandaloneEdits(form)
        if rect.left >= form_mid
    ]


def FillStandaloneEditByIndex(edits, index, value, label, log_fn, index_name):
    if index >= len(edits):
        log_fn(f"  |  `- FAIL {label} {index_name}[{index}] not found", RED_ERR)
        return False

    edit = edits[index][0]
    log_fn(f"  |  `- using {index_name}[{index}]", AMBER)
    return FillEditControl(edit, value, label, log_fn)


def FillEditByIndex(form, index, value, label, log_fn):
    edits = GetVisibleControls(form, "TEdit")
    if index >= len(edits):
        log_fn(f"  |  `- FAIL {label} edit index={index} not found", RED_ERR)
        return False

    edit = edits[index][0]
    log_fn(f"  | [{label}] = '{value}'", BLUE)
    log_fn(f"  |  `- target edit[{index}] {ControlDebugInfo(edit)}", TEXT_SEC)
    try:
        edit.set_edit_text(value)
        log_fn("  |  `- OK Set via set_edit_text()", GREEN)
        return True
    except Exception as e:
        log_fn(f"  |  `- WARN set_edit_text failed: {e}; trying clipboard", AMBER)
        try:
            import pyperclip
            pyperclip.copy(value)
            edit.click_input()
            time.sleep(0.1)
            edit.type_keys("^a^v")
            log_fn("  |  `- OK Set via clipboard", GREEN)
            return True
        except Exception as e2:
            log_fn(f"  |  `- FAIL {e2}", RED_ERR)
            return False


def FillScrapWindow(form, sn, cfg, log_fn, status_fn):
    location_code = cfg.get("LOCATION_CODE", "")
    memo_template = cfg.get("MEMO_TEMPLATE", "{sn}")
    try:
        memo = memo_template.format(sn=sn)
    except Exception:
        memo = f"{memo_template}{sn}"

    combos = GetVisibleControls(form, "TComboBox")
    edits = GetVisibleControls(form, "TEdit")
    standalone_edits = GetStandaloneEdits(form)
    left_standalone_edits = GetLeftStandaloneEdits(form)
    right_standalone_edits = GetRightStandaloneEdits(form)
    form_rect = form.rectangle()
    form_mid = form_rect.left + form_rect.width() // 2
    left_edits = [(c, r) for c, r in edits if r.left < form_mid]
    right_edits = [(c, r) for c, r in edits if r.left >= form_mid]

    if DEBUG_CONTROL_LOGS:
        log_fn("  + Scrap control map by sorted position:", BLUE)
        for idx, (combo, _r) in enumerate(combos):
            log_fn(f"  | combo[{idx}] {ControlDebugInfo(combo)}", TEXT_SEC)
        for idx, (edit, _r) in enumerate(left_edits):
            log_fn(f"  | left_edit[{idx}] {ControlDebugInfo(edit)}", TEXT_SEC)
        for idx, (edit, _r) in enumerate(right_edits):
            log_fn(f"  | right_edit[{idx}] {ControlDebugInfo(edit)}", TEXT_SEC)
        for idx, (edit, _r) in enumerate(standalone_edits):
            log_fn(f"  | standalone_edit[{idx}] {ControlDebugInfo(edit)}", TEXT_SEC)
        for idx, (edit, _r) in enumerate(left_standalone_edits):
            log_fn(f"  | left_standalone_edit[{idx}] {ControlDebugInfo(edit)}", TEXT_SEC)
        for idx, (edit, _r) in enumerate(right_standalone_edits):
            log_fn(f"  | right_standalone_edit[{idx}] {ControlDebugInfo(edit)}", TEXT_SEC)

    status_fn("FILL SCRAP", AMBER)
    FillLeftComboByIndex(form, 0, cfg.get("SCRAP_CODE", ""), "Scrap Code", log_fn)
    blank_left_standalone_edits = [
        (edit, rect) for edit, rect in left_standalone_edits
        if not _control_text(edit)
    ]
    FillStandaloneEditByIndex(
        blank_left_standalone_edits,
        1,
        location_code,
        "Location Code",
        log_fn,
        "blank_left_standalone_edit",
    )

    status_fn("FILL DUTY", AMBER)
    FillLeftComboByIndex(form, 4, cfg.get("DUTY_CODE", ""), "Duty Code", log_fn)
    time.sleep(0.4)
    FillLeftReasonCombo(form, 5, cfg.get("REASON_CODE", ""), log_fn)
    FillLeftComboByIndex(form, 6, cfg.get("HANDLING", ""), "Handling", log_fn)
    FillLeftComboByIndex(form, 7, cfg.get("DUTY_DEPARTMENT", ""), "Duty Department", log_fn)

    status_fn("FILL COST", AMBER)
    FillLeftComboByIndex(form, 2, cfg.get("COST_CENTER", ""), "Cost Center", log_fn)

    status_fn("FILL MEMO", AMBER)
    FillMemoSmart(form, memo, log_fn)


def ClickCancel(form):
    for title in ("Cancel", "CANCEL"):
        for class_name in ("TBitBtn", "TButton", "Button"):
            btn = form.child_window(title=title, class_name=class_name)
            if btn.exists():
                btn.click_input()
                time.sleep(0.3)
                return
    raise RuntimeError("Cancel button not found")


def ClickOK(form):
    for title in ("OK", "Ok", "ok"):
        for class_name in ("TBitBtn", "TButton", "Button"):
            btn = form.child_window(title=title, class_name=class_name)
            if btn.exists():
                btn.click_input()
                time.sleep(0.3)
                return
    raise RuntimeError("OK button not found")


def _win32_child_texts(hwnd):
    texts = []

    def _enum(child_hwnd, _param):
        try:
            text = win32gui.GetWindowText(child_hwnd).strip()
            if text:
                texts.append(text)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _enum, None)
    except Exception:
        pass
    return texts


def _click_win32_ok(hwnd):
    ok_hwnds = []

    def _enum(child_hwnd, _param):
        try:
            text = win32gui.GetWindowText(child_hwnd).strip().replace("&", "").lower()
            cls = win32gui.GetClassName(child_hwnd)
            if text == "ok" and cls.lower() == "button":
                ok_hwnds.append(child_hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumChildWindows(hwnd, _enum, None)
    if not ok_hwnds:
        return False

    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    win32gui.SendMessage(ok_hwnds[0], win32con.BM_CLICK, 0, 0)
    time.sleep(0.3)
    return True


def _click_dialog_ok(win):
    try:
        ClickOK(win)
        return True
    except Exception:
        pass
    try:
        return _click_win32_ok(win.handle)
    except Exception:
        return False


def ClickScrapSuccessOK(log_fn, timeout=8):
    expected_text = "scrap the production successfully"
    start = time.time()
    while time.time() - start < timeout:
        windows = []
        seen = set()

        for win in Desktop(backend="win32").windows(visible_only=True):
            try:
                windows.append(win)
                seen.add(win.handle)
            except Exception:
                pass

        def _enum(hwnd, _param):
            try:
                if win32gui.IsWindowVisible(hwnd) and hwnd not in seen:
                    windows.append(Desktop(backend="win32").window(handle=hwnd))
                    seen.add(hwnd)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(_enum, None)

        for win in windows:
            try:
                title = win.window_text().strip()
                title_lower = title.lower()

                texts = [title_lower]
                texts.extend(text.lower() for text in _win32_child_texts(win.handle))
                for child in win.descendants():
                    try:
                        text = child.window_text().strip()
                        if text:
                            texts.append(text.lower())
                    except Exception:
                        pass

                is_information = "information" in title_lower
                has_success_text = any(expected_text in text for text in texts)
                has_ok_button = any(
                    text.strip().replace("&", "").lower() == "ok"
                    for text in _win32_child_texts(win.handle)
                )
                if not has_ok_button or not (has_success_text or is_information):
                    continue

                if not _click_dialog_ok(win):
                    continue
                log_fn("  `- OK Scrap success information confirmed", GREEN)
                return True
            except Exception:
                pass
        time.sleep(0.2)

    log_fn("  `- WARN Scrap success information dialog not found", AMBER)
    DumpTopWindows(log_fn)
    return False


def _looks_like_privilege_title(title):
    text = (title or "").strip().lower()
    return "previlege control" in text


def _visible_window_titles():
    titles = []
    for w in Desktop(backend="win32").windows(visible_only=True):
        try:
            title = w.window_text().strip()
            if title:
                titles.append(title)
        except Exception:
            pass
    return titles


def DumpTopWindows(log_fn):
    log_fn("  + Top-level windows:", BLUE)

    def _enum(hwnd, _param):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd).strip()
            cls = win32gui.GetClassName(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            if title or "control" in cls.lower() or "repair" in cls.lower():
                log_fn(
                    f"  | hwnd={hwnd} class={cls} title='{title}' rect={rect}",
                    TEXT_SEC,
                )
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_enum, None)


def _find_privilege_hwnd_by_win32():
    matches = []

    def _enum(hwnd, _param):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd).strip()
            if _looks_like_privilege_title(title):
                matches.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_enum, None)
    return matches[0] if matches else None


def _find_privilege_child_hwnd(parent_hwnd):
    matches = []

    def _enum(hwnd, _param):
        try:
            title = win32gui.GetWindowText(hwnd).strip()
            if _looks_like_privilege_title(title):
                matches.append(hwnd)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(parent_hwnd, _enum, None)
    except Exception:
        pass
    return matches[0] if matches else None


def _find_privilege_child_by_pywinauto(parent_form):
    try:
        for c in parent_form.descendants():
            try:
                if _looks_like_privilege_title(_control_text(c)):
                    return c
                if _looks_like_privilege_title(c.window_text()):
                    return c
            except Exception:
                pass
    except Exception:
        pass
    return None


def FindPrivilegeEditFields(container):
    labels = {}
    edits = []
    try:
        descendants = container.descendants()
    except Exception:
        descendants = []

    for c in descendants:
        try:
            if not c.is_visible():
                continue
            text = _norm_label(_control_text(c))
            rect = c.rectangle()
            if text in ("emp", "password"):
                labels[text] = rect
            if _class_matches(c.class_name(), "TEdit"):
                edits.append((c, rect))
        except Exception:
            pass

    found = []
    for label_name in ("emp", "password"):
        lr = labels.get(label_name)
        if not lr:
            return []
        label_y = (lr.top + lr.bottom) // 2
        best = None
        best_score = None
        for edit, er in edits:
            edit_y = (er.top + er.bottom) // 2
            if er.left < lr.right - 5:
                continue
            dy = abs(edit_y - label_y)
            if dy > 28:
                continue
            dx = er.left - lr.right
            score = dy * 1000 + dx
            if best_score is None or score < best_score:
                best = edit
                best_score = score
        if not best:
            return []
        found.append(best)

    return found


def FindEditChildrenByWin32(hwnd):
    edits = []

    def _enum(child_hwnd, _param):
        try:
            if not win32gui.IsWindowVisible(child_hwnd):
                return True
            cls = win32gui.GetClassName(child_hwnd)
            if "edit" in cls.lower():
                rect = win32gui.GetWindowRect(child_hwnd)
                ctrl = Desktop(backend="win32").window(handle=child_hwnd)
                edits.append((ctrl, rect))
        except Exception:
            pass
        return True

    try:
        win32gui.EnumChildWindows(hwnd, _enum, None)
    except Exception:
        pass

    edits.sort(key=lambda item: (item[1][1], item[1][0]))
    return [ctrl for ctrl, _rect in edits]


def FillPrivilegeByCoordinates(pform, emp, password, log_fn):
    rect = pform.rectangle()
    width = rect.width()
    height = rect.height()
    emp_pos = (int(width * 0.60), int(height * 0.34))
    password_pos = (int(width * 0.60), int(height * 0.55))

    log_fn("  |  `- WARN using coordinate fallback for privilege fields", AMBER)

    pform.click_input(coords=emp_pos)
    time.sleep(0.1)
    pform.type_keys("^a{BACKSPACE}")
    pform.type_keys(str(emp), with_spaces=True)

    pform.click_input(coords=password_pos)
    time.sleep(0.1)
    pform.type_keys("^a{BACKSPACE}")
    pform.type_keys(str(password), with_spaces=True)
    log_fn("  |  `- OK Filled privilege fields by coordinates", GREEN)
    return True


def _parent_has_privilege_fields(parent_form):
    return len(FindPrivilegeEditFields(parent_form)) >= 2


def _window_has_privilege_fields(win):
    try:
        texts = []
        edit_count = 0
        for c in win.descendants():
            try:
                if not c.is_visible():
                    continue
                text = _control_text(c).lower()
                if text:
                    texts.append(text)
                if _class_matches(c.class_name(), "TEdit"):
                    edit_count += 1
            except Exception:
                pass
        joined = " ".join(texts)
        return "emp" in joined and "password" in joined and edit_count >= 2
    except Exception:
        return False


def _find_privilege_window_by_fields():
    for w in Desktop(backend="win32").windows(visible_only=True):
        try:
            if _window_has_privilege_fields(w):
                return w
        except Exception:
            pass
    return None


def WaitForPrivilegeControl(timeout=12, log_fn=None, parent_form=None):
    start = time.time()
    while time.time() - start < timeout:
        if parent_form is not None:
            if _parent_has_privilege_fields(parent_form):
                return parent_form
            child = _find_privilege_child_by_pywinauto(parent_form)
            if child:
                return child
            try:
                hwnd = _find_privilege_child_hwnd(parent_form.handle)
                if hwnd:
                    return Desktop(backend="win32").window(handle=hwnd)
            except Exception:
                pass

        for w in Desktop(backend="win32").windows(visible_only=True):
            try:
                if _looks_like_privilege_title(w.window_text()):
                    return w
            except Exception:
                pass
        hwnd = _find_privilege_hwnd_by_win32()
        if hwnd:
            return Desktop(backend="win32").window(handle=hwnd)
        field_win = _find_privilege_window_by_fields()
        if field_win:
            return field_win
        time.sleep(0.2)
    if log_fn:
        DumpTopWindows(log_fn)
    titles = ", ".join(_visible_window_titles()[:12])
    hwnd_titles = []
    def _enum_titles(hwnd, _param):
        try:
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd).strip()
                if title:
                    hwnd_titles.append(title)
        except Exception:
            pass
        return True
    win32gui.EnumWindows(_enum_titles, None)
    raise RuntimeError(
        f"Timeout {timeout}s: Privilege Control did not appear. "
        f"Visible windows: {titles}. Win32 titles: {', '.join(hwnd_titles[:16])}"
    )


def FillPrivilegeControl(cfg, log_fn, status_fn, parent_form=None):
    status_fn("PRIVILEGE", AMBER)
    log_fn("  | Waiting for Privilege Control...", BLUE)
    privilege_win = WaitForPrivilegeControl(timeout=12, log_fn=log_fn, parent_form=parent_form)
    log_fn(f"  | Privilege window found: '{privilege_win.window_text()}'", GREEN)
    papp = Application(backend="win32").connect(handle=privilege_win.handle)
    pform = papp.window(handle=privilege_win.handle)
    pform.set_focus()
    time.sleep(0.2)

    emp = cfg.get("PRIVILEGE_EMP", "")
    password = cfg.get("PRIVILEGE_PASSWORD", "")
    fields = FindPrivilegeEditFields(pform)
    if len(fields) < 2 and parent_form is not None:
        fields = FindPrivilegeEditFields(parent_form)
    if len(fields) < 2:
        fields = FindEditChildrenByWin32(privilege_win.handle)
    if len(fields) < 2 and parent_form is not None:
        fields = FindEditChildrenByWin32(parent_form.handle)
    if len(fields) < 2:
        filled = FillPrivilegeByCoordinates(pform, emp, password, log_fn)
    else:
        for idx, edit in enumerate(fields):
            log_fn(f"  | privilege_edit[{idx}] {ControlDebugInfo(edit)}", TEXT_SEC)

        emp_filled = FillEditControl(fields[0], emp, "Privilege Emp", log_fn)
        password_filled = FillEditControl(fields[1], password, "Privilege Password", log_fn)
        filled = emp_filled and password_filled

    if not filled:
        raise RuntimeError("Privilege employee ID or password could not be filled")

    log_fn("  | Clicking Privilege OK", BLUE)
    ClickOK(pform)
    time.sleep(0.5)


def ClickChange(main_form, log_fn):
    for title in ("Change", "CHANGE"):
        btn = main_form.child_window(title=title, class_name="TBitBtn")
        if btn.exists():
            btn.click_input()
            time.sleep(0.5)
            log_fn("  `- OK Change clicked - ready for next SN", GREEN)
            return True
    log_fn("  `- WARN Change button not found", AMBER)
    return False


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


def GetFirstRedErrorCodeScrap(main_form, cfg, sn, log_fn, status_fn):
    code = None
    found_red_row = False

    status_fn("FINDING GRID", AMBER)
    log_fn("  + Step 1/7 - Finding Error Code Grid...", BLUE)
    target_grid = FindErrorCodeDBGrid(main_form)
    if not target_grid:
        log_fn("  `- FAIL Error Code Grid not found", RED_ERR)
        return None, False, False
    log_fn(f"  `- OK Grid found at {target_grid.rectangle()}", GREEN)

    grid_rect = target_grid.rectangle()

    status_fn("CAPTURING", AMBER)
    log_fn("  + Step 2/7 - Capturing window image...", BLUE)
    try:
        img = GrabWindow(main_form.handle, rect=grid_rect)
    except Exception as e:
        log_fn(f"  `- FAIL Capture failed: {e}", RED_ERR)
        return None, False, False
    pixels = img.load()
    width, height = img.size
    log_fn(f"  `- OK Captured {width}x{height}px (saved: debug_grid.png)", GREEN)

    status_fn("SCANNING", AMBER)
    log_fn("  + Step 3/7 - Scanning for red highlighted row...", BLUE)
    first_red_y = None
    max_pct = 0.0
    for y in range(height):
        red_count = sum(1 for x in range(width) if is_red_bg(*pixels[x, y]))
        pct = red_count / width * 100
        max_pct = max(max_pct, pct)
        if pct > 5.0 and first_red_y is None:
            first_red_y = grid_rect.top + y
            found_red_row = True
            log_fn(f"  `- OK Red row at screen Y={first_red_y} ({pct:.1f}%)", GREEN)

    if not found_red_row:
        log_fn(f"  `- OK No red rows found (max red={max_pct:.1f}%) - PASS", GREEN)
        status_fn("CLICK CHANGE", AMBER)
        ClickChange(main_form, log_fn)
        return None, False, False

    status_fn("SELECTING", AMBER)
    log_fn("  + Step 4/7 - Clicking red row...", BLUE)
    main_rect = main_form.rectangle()
    click_x = grid_rect.left + (grid_rect.right - grid_rect.left) // 2
    main_form.click_input(coords=(click_x - main_rect.left, first_red_y - main_rect.top))
    time.sleep(0.5)
    log_fn(f"  `- OK Clicked at X={click_x}, Y={first_red_y}", GREEN)

    status_fn("READING CODE", AMBER)
    log_fn("  + Step 5/7 - Reading error code from TDBEdit...", BLUE)
    error_code_edit = FindErrorCodeEdit(main_form)
    if error_code_edit:
        raw_text = error_code_edit.texts()[0].strip()
        if raw_text:
            code = raw_text
            log_fn(f"  `- OK Error code: '{code}'", GREEN)
        else:
            log_fn("  `- WARN TDBEdit found but empty", AMBER)
    else:
        log_fn("  `- FAIL TDBEdit not found", RED_ERR)

    status_fn("CLICK SCRAP", AMBER)
    log_fn("  + Step 6/7 - Clicking Scrap button...", BLUE)
    try:
        scrap_btn = main_form.child_window(title="Scrap", class_name="TBitBtn")
        if not scrap_btn.exists():
            scrap_btn = main_form.child_window(title="SCRAP", class_name="TBitBtn")
        if not scrap_btn.exists():
            log_fn("  `- FAIL Scrap button not found", RED_ERR)
            return code, found_red_row, False

        scrap_btn.click()
        log_fn("  `- OK Scrap clicked - waiting for Repair Window...", GREEN)
        repair_win = WaitForRepairWindow(timeout=5)

        status_fn("FILL SCRAP", AMBER)
        log_fn("  + Step 7/7 - Filling Scrap Window...", BLUE)
        rapp = Application(backend="win32").connect(handle=repair_win.handle)
        rform = rapp.window(handle=repair_win.handle)
        rform.set_focus()
        WaitForScrapControls(rform, log_fn, timeout=5)
        if DEBUG_CONTROL_LOGS:
            DumpScrapWindowControls(rform, log_fn)
        FillScrapWindow(rform, sn, cfg, log_fn, status_fn)

        status_fn("SUBMIT", AMBER)
        log_fn("  | Clicking Scrap OK to submit", BLUE)
        rform_handle = rform.handle
        ClickOK(rform)
        FillPrivilegeControl(cfg, log_fn, status_fn, parent_form=rform)

        status_fn("WAIT CLOSE", AMBER)
        log_fn("  | Waiting for Scrap form to close after submit", BLUE)
        if WaitForWindowGone(rform_handle, timeout=5):
            log_fn("  `- OK Scrap form closed after submit", GREEN)
        else:
            log_fn("  `- WARN Scrap form still visible after submit", AMBER)

        status_fn("CONFIRM SUCCESS", AMBER)
        log_fn("  | Waiting for Scrap success information OK", BLUE)
        if not ClickScrapSuccessOK(log_fn):
            return code, found_red_row, False

        status_fn("CLICK CHANGE", AMBER)
        ClickChange(main_form, log_fn)

    except Exception as e:
        log_fn(f"  `- FAIL Scrap flow error: {e}", RED_ERR)
        return code, found_red_row, False

    return code, found_red_row, True


# ============================================================
#  MAIN PROCESS
# ============================================================
def RunRepairProcess(sn, log_fn, status_fn, cfg=None):
    try:
        cfg = cfg or LoadConfig()
        log_fn("· Scrap config loaded:", BLUE)
        log_fn(f"  Scrap Code : {cfg.get('SCRAP_CODE', '')}", TEXT_SEC)
        log_fn(f"  Location   : {cfg.get('LOCATION_CODE', '')}", TEXT_SEC)
        log_fn(f"  Duty Code  : {cfg.get('DUTY_CODE', '')}", TEXT_SEC)
        log_fn(f"  Reason     : {cfg.get('REASON_CODE', '')}", TEXT_SEC)
        log_fn(f"  Handling   : {cfg.get('HANDLING', '')}", TEXT_SEC)
        log_fn(f"  Department : {cfg.get('DUTY_DEPARTMENT', '')}", TEXT_SEC)
        log_fn(f"  Cost Center: {cfg.get('COST_CENTER', '')}", TEXT_SEC)

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
        code, red_row_exists, scrap_completed = GetFirstRedErrorCodeScrap(
            repair_form, cfg, sn, log_fn, status_fn
        )

        if scrap_completed:
            log_fn(f"  RESULT - Scrap form completed: {code or 'code unavailable'}", GREEN)
            status_fn("COMPLETE", GREEN)
            return {
                "completed": True,
                "code": code,
                "message": "Scrap form completed",
            }
        if red_row_exists:
            log_fn("  RESULT - Scrap form was not completed", RED_ERR)
            status_fn("FAILED", RED_ERR)
            return {
                "completed": False,
                "code": code,
                "message": "Scrap form was not completed",
            }

        log_fn("  RESULT - No red repair row found; scrap form was not completed", RED_ERR)
        status_fn("FAILED", RED_ERR)
        return {
            "completed": False,
            "code": code,
            "message": "No red repair row found",
        }

    except Exception as e:
        log_fn(f"✗ EXCEPTION : {e}", RED_ERR)
        status_fn("FAILED", RED_ERR)
        return {
            "completed": False,
            "code": None,
            "message": str(e),
        }


# ============================================================
#  GUI
# ============================================================
class RepairGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Repair Automation")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(False, False)
        self.root.geometry("720x880")
        self.sn_rows = []
        self.log_lock = threading.Lock()
        os.makedirs(LOG_DIR, exist_ok=True)
        log_name = datetime.now().strftime("repair_debug_%Y%m%d_%H%M%S.txt")
        self.log_path = os.path.join(LOG_DIR, log_name)

        self._build_fonts()
        self._build_ui()
        self._set_ready()
        self._log(f"Debug log file: {self.log_path}", AMBER)
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
        self.import_btn = tk.Button(
            br, text="IMPORT EXCEL", font=self.f_btn,
            bg=BG_PANEL, fg=TEXT_PRI,
            activebackground=BORDER, activeforeground=TEXT_PRI,
            relief="flat", bd=0, cursor="hand2",
            command=self._import_excel, padx=18, pady=8)
        self.import_btn.pack(side="left", padx=(10, 0))
        self.run_btn = tk.Button(
            br, text="▶  RUN", font=self.f_btn,
            bg=AMBER, fg=BG_DARK,
            activebackground=AMBER_DIM, activeforeground=TEXT_PRI,
            relief="flat", bd=0, cursor="hand2",
            command=lambda: self._on_scan(None), padx=18, pady=8)
        self.run_btn.pack(side="right")

        tk.Label(self.root, text="SN BATCH STATUS", font=self.f_badge,
                 bg=BG_DARK, fg=TEXT_SEC, anchor="w").pack(
                     fill="x", padx=20, pady=(14, 2))
        tf = tk.Frame(self.root, bg=BG_PANEL,
                      highlightbackground=BORDER, highlightthickness=1)
        tf.pack(fill="x", padx=20)
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Treeview", background=BG_PANEL, foreground=TEXT_PRI,
                        fieldbackground=BG_PANEL, rowheight=24, borderwidth=0)
        style.configure("Treeview.Heading", background=BG_INPUT,
                        foreground=TEXT_PRI, borderwidth=0)
        self.sn_tree = ttk.Treeview(
            tf, columns=("idx", "sn", "state", "result"),
            show="headings", height=6)
        self.sn_tree.heading("idx", text="#")
        self.sn_tree.heading("sn", text="SN")
        self.sn_tree.heading("state", text="STATE")
        self.sn_tree.heading("result", text="RESULT")
        self.sn_tree.column("idx", width=44, anchor="center", stretch=False)
        self.sn_tree.column("sn", width=210, anchor="w")
        self.sn_tree.column("state", width=190, anchor="w")
        self.sn_tree.column("result", width=210, anchor="w")
        self.sn_tree.pack(fill="x")

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
        self.import_btn.config(state="normal")

    def _set_busy(self):
        self.running = True
        self.sn_entry.config(state="disabled")
        self.run_btn.config(state="disabled", bg=BORDER)
        self.reset_btn.config(state="disabled")
        self.import_btn.config(state="disabled")

    def _log(self, message, color=None):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {message}\n"
        try:
            with self.log_lock:
                with open(self.log_path, "a", encoding="utf-8") as f:
                    f.write(line)
        except Exception:
            pass

        def _write():
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

    def _import_excel(self):
        if self.running:
            return
        path = filedialog.askopenfilename(
            title="Import SN Excel",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        try:
            serials = ReadSerialsFromExcel(path)
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
            self._log(f"Import failed: {e}", RED_ERR)
            return

        if not serials:
            messagebox.showwarning("No SN found", "No serial numbers were found.")
            return

        self.sn_rows = []
        for item in self.sn_tree.get_children():
            self.sn_tree.delete(item)
        for idx, sn in enumerate(serials, start=1):
            item_id = self.sn_tree.insert("", "end", values=(idx, sn, "PENDING", ""))
            self.sn_rows.append({"sn": sn, "item_id": item_id})

        self._set_status("IMPORTED", GREEN)
        self._set_result(f"{len(serials)} SN loaded", GREEN)
        self._log(f"Imported {len(serials)} SN(s) from {os.path.basename(path)}", GREEN)

    def _set_sn_state(self, sn, state, result=""):
        def _write():
            for idx, row in enumerate(self.sn_rows, start=1):
                if row["sn"] == sn:
                    values = list(self.sn_tree.item(row["item_id"], "values"))
                    current_result = values[3] if len(values) > 3 else ""
                    self.sn_tree.item(
                        row["item_id"],
                        values=(idx, sn, state, result or current_result),
                    )
                    self.sn_tree.see(row["item_id"])
                    break
        self.root.after(0, _write)

    def _show_context_dialog(self, on_submit):
        cfg = LoadConfig()
        dialog = tk.Toplevel(self.root)
        dialog.title("Scrap Context")
        dialog.configure(bg=BG_DARK)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.attributes("-topmost", True)

        fields = [
            ("SCRAP_CODE", "Scrap Code"),
            ("LOCATION_CODE", "Location Code"),
            ("DUTY_CODE", "Duty Code"),
            ("REASON_CODE", "Reason Code"),
            ("HANDLING", "Handling"),
            ("DUTY_DEPARTMENT", "Duty Department"),
            ("COST_CENTER", "Cost Center"),
            ("MEMO_TEMPLATE", "Memo Template"),
            ("PRIVILEGE_EMP", "Privilege Emp"),
            ("PRIVILEGE_PASSWORD", "Privilege Password"),
        ]
        values = {}
        result = {"cfg": None}

        tk.Label(dialog, text="SCRAP FORM CONTEXT", font=self.f_title,
                 bg=BG_DARK, fg=AMBER).grid(
                     row=0, column=0, columnspan=2, sticky="w",
                     padx=16, pady=(14, 8))
        tk.Label(dialog, text="Memo supports {sn}",
                 font=self.f_badge, bg=BG_DARK, fg=TEXT_SEC).grid(
                     row=1, column=0, columnspan=2, sticky="w",
                     padx=16, pady=(0, 10))

        for row, (key, label) in enumerate(fields, start=2):
            tk.Label(dialog, text=label, font=self.f_badge,
                     bg=BG_DARK, fg=TEXT_SEC).grid(
                         row=row, column=0, sticky="w", padx=(16, 8), pady=4)
            var = tk.StringVar(value=str(cfg.get(key, "")))
            values[key] = var
            ent = tk.Entry(dialog, textvariable=var, font=self.f_label,
                           bg=BG_INPUT, fg=TEXT_PRI,
                           insertbackground=AMBER, relief="flat", width=46)
            ent.grid(row=row, column=1, sticky="ew", padx=(0, 16), pady=4, ipady=5)

        test_row = len(fields) + 2

        btns = tk.Frame(dialog, bg=BG_DARK)
        btns.grid(row=test_row, column=0, columnspan=2,
                  sticky="e", padx=16, pady=(10, 16))

        def submit():
            new_cfg = dict(cfg)
            for key, var in values.items():
                new_cfg[key] = var.get().strip()
            new_cfg["MODE"] = "SCRAP"
            result["cfg"] = new_cfg
            dialog.destroy()
            self.root.after(0, lambda: on_submit(new_cfg))

        def cancel():
            result["cfg"] = None
            dialog.destroy()

        tk.Button(btns, text="CANCEL", font=self.f_btn,
                  bg=BG_PANEL, fg=TEXT_PRI, relief="flat", bd=0,
                  command=cancel, padx=18, pady=8).pack(side="right", padx=(8, 0))
        tk.Button(btns, text="RUN", font=self.f_btn,
                  bg=AMBER, fg=BG_DARK, relief="flat", bd=0,
                  command=submit, padx=18, pady=8).pack(side="right")

        dialog.bind("<Escape>", lambda e: cancel())
        dialog.bind("<Return>", lambda e: submit())
        dialog.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_rooty() + 80
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dialog.lift()
        dialog.focus_force()
        dialog.after(250, lambda: dialog.attributes("-topmost", False))
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        return dialog

    def _on_scan_single_unused(self, event):
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
            outcome = RunRepairProcess(
                sn=sn,
                log_fn=self._log,
                status_fn=self._set_status,
            )
            if outcome["completed"]:
                result = outcome["code"] or outcome["message"]
                self.root.after(0, lambda r=result: self._set_result(r, GREEN))
            else:
                self.root.after(
                    0,
                    lambda m=outcome["message"]: self._set_result(m, RED_ERR),
                )
            self.root.after(0, self._after_process)

        threading.Thread(target=worker, daemon=True).start()

    def _on_scan(self, event):
        if self.running:
            return
        sns = [row["sn"] for row in self.sn_rows]
        if not sns:
            sn = self.sn_var.get().strip()
            if sn:
                item_id = self.sn_tree.insert("", "end", values=(1, sn, "PENDING", ""))
                self.sn_rows = [{"sn": sn, "item_id": item_id}]
                sns = [sn]

        if not sns:
            self._log("No serial number entered", RED_ERR)
            return

        self._show_context_dialog(lambda run_cfg: self._start_batch(sns, run_cfg))

    def _start_batch(self, sns, run_cfg):
        self._set_busy()
        self._set_status("RUNNING...", AMBER)
        self._set_result("-", TEXT_SEC)

        def worker():
            total = len(sns)
            done = 0
            failed = 0
            for idx, sn in enumerate(sns, start=1):
                self._set_sn_state(sn, "START")
                self._set_status(f"SN {idx}/{total}", AMBER)
                self._log(f"Batch {idx}/{total}: {sn}", TEXT_PRI)

                def status_for_sn(text, color=GREEN, sn=sn):
                    self._set_status(f"SN {idx}/{total}: {text}", color)
                    self._set_sn_state(sn, text)

                try:
                    outcome = RunRepairProcess(
                        sn=sn,
                        log_fn=self._log,
                        status_fn=status_for_sn,
                        cfg=run_cfg,
                    )
                    if outcome["completed"]:
                        done += 1
                        result = outcome["code"] or outcome["message"]
                        self._set_sn_state(sn, "COMPLETE", result)
                        self.root.after(0, lambda r=result: self._set_result(r, GREEN))
                    else:
                        failed += 1
                        message = outcome["message"]
                        self._set_sn_state(sn, "FAILED", message)
                        self.root.after(0, lambda m=message: self._set_result(m, RED_ERR))
                except Exception as e:
                    failed += 1
                    self._set_sn_state(sn, "FAILED", str(e))
                    self._log(f"SN {sn} failed: {e}", RED_ERR)

            summary_color = GREEN if failed == 0 else RED_ERR
            summary = f"COMPLETE {done}/{total}" if failed == 0 else f"FAILED {failed}/{total}"
            self._set_status(summary, summary_color)
            self._log(f"Batch complete: {done} completed, {failed} failed", summary_color)
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
        self.sn_rows = []
        for item in self.sn_tree.get_children():
            self.sn_tree.delete(item)
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
