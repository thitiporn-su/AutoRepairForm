import tkinter as tk
from tkinter import font as tkfont
import threading
import time
import ctypes
import PIL.ImageGrab
import numpy as np
import win32gui
import win32ui
from PIL import ImageGrab
from PIL import Image
from datetime import datetime
from pywinauto import Application, Desktop, mouse, keyboard

import ctypes
import ReadData

data = ReadData.GetFormData("sn")



# ============================================================
#  COLOR PALETTE  — dark industrial / amber accent
# ============================================================
BG_DARK    = "#0e0f11"
BG_PANEL   = "#16181c"
BG_INPUT   = "#1e2128"
BORDER     = "#2a2d35"
AMBER      = "#f5a623"
AMBER_DIM  = "#7a5212"
GREEN      = "#3ddc84"
RED_ERR    = "#ff4f4f"
TEXT_PRI   = "#e8eaf0"
TEXT_SEC   = "#6b7280"
TEXT_MONO  = "#a8b0c0"
CYAN       = "#38bdf8"


import sys

import struct

import warnings
 
def CheckPythonBitness():

    """

    ตรวจสอบ Python bitness และแจ้งเตือนถ้าไม่ตรงกับ target app

    suppress warning ของ pywinauto ไปในตัว

    """

    # suppress pywinauto warning เสมอ

    warnings.filterwarnings(

        "ignore",

        category=UserWarning,

        module="pywinauto"

    )
 
    bits = struct.calcsize("P") * 8  # 32 หรือ 64

    if bits == 64:

        print(

            f"[WARN] Running 64-bit Python — automating 32-bit app may have issues.\n"

            f"       Recommended: install Python 32-bit from python.org\n"

            f"       Current: {sys.executable}"

        )

    else:

        print(f"[INFO] Python {bits}-bit — OK")
 
# ============================================================
#  DPI / COLOR HELPERS
# ============================================================

def SetDPIAware():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

def GetScaleFactor():
    try:
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return dpi / 96.0
    except Exception:
        return 1.0

def GrabWithDPI(rect):
    scale = GetScaleFactor()
    return PIL.ImageGrab.grab(bbox=(
        int(rect.left   * scale),
        int(rect.top    * scale),
        int(rect.right  * scale),
        int(rect.bottom * scale),
    ))

def is_red_bg(r, g, b):
    return r > 200 and g < 60 and b < 60


# ============================================================
#  REPAIR LOGIC
# ============================================================

def WaitForMainForm(timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        all_windows = Desktop(backend="win32").windows(title_re=r"^Repair-Rev")
        for w in all_windows:
            if w.class_name() == "TfrmMain":
                return w
        time.sleep(0.5)
    raise RuntimeError(f"Timeout {timeout}s: TfrmMain did not appear")


def WaitForRepairWindow(timeout=10):
    start = time.time()
    while time.time() - start < timeout:
        wins = Desktop(backend="uia").windows(title_re=r"^Repair Window")
        if wins:
            return wins[0]
        time.sleep(0.5)
    raise RuntimeError(f"Timeout {timeout}s: Repair Window did not appear")


def GrabWindow(hwnd, rect=None):
    client_rect = win32gui.GetClientRect(hwnd)
    client_left, client_top = win32gui.ClientToScreen(hwnd, (0, 0))
    win_w = client_rect[2]
    win_h = client_rect[3]

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
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
        bmp_bits,
        "raw",
        "BGRX",
        0,
        1,
    )

    win32gui.DeleteObject(bmp.GetHandle())
    save_dc.DeleteDC()
    mfc_dc.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwnd_dc)

    if rect is not None:
        crop_l = rect.left - client_left
        crop_t = rect.top - client_top
        crop_r = rect.right - client_left
        crop_b = rect.bottom - client_top
        img = img.crop((crop_l, crop_t, crop_r, crop_b))

    img.save("debug_grid.png")
    return img

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
    anchor_btn = None
    for child in main_form.descendants():
        try:
            if child.class_name() == "TBitBtn" and child.texts()[0] in ("New", "Remove"):
                anchor_btn = child
                break
        except Exception:
            pass
    if anchor_btn:
        btn_rect = anchor_btn.rectangle()
        for grid in all_grids:
            r = grid.rectangle()
            if r.bottom <= btn_rect.top and abs(r.left - btn_rect.left) < 50:
                return grid
    return min(all_grids, key=lambda g: g.rectangle().left)

def FindErrorCodeEdit(main_form):
    all_dbedits = []
    for child in main_form.descendants():
        try:
            if child.class_name() == "TDBEdit":
                all_dbedits.append(child)
        except Exception:
            pass
    for edit in all_dbedits:
        try:
            text = edit.texts()[0].strip()
            if text.isdigit() and len(text) == 12:
                return edit
        except Exception:
            pass
    if len(all_dbedits) >= 2:
        return all_dbedits[1]
    return None

def ClickAddRepairButton(main_form, log_fn=None):
    """
    ใช้ pywinauto ค้นหาและคลิกปุ่ม Add โดยตรง 
    ไม่ต้องสแกนสีพิกเซล เพื่อความแม่นยำสูงสุด
    """
    try:
        # 1. ค้นหาปุ่มที่มีคำว่า "Add" (จากรูปที่คุณส่งมา ปุ่มชื่อ Add และมีเครื่องหมายบวก)
        # เราหาจาก descendants ที่เป็น Class TBitBtn (มาตรฐานของ Delphi)
        btn_new = None
        for ctrl in main_form.descendants():
            if ctrl.class_name() == "TBitBtn" and "Add" in ctrl.window_text():
                btn_new = ctrl
                break
        
        if btn_new:
            # ใช้ click_input() เพื่อจำลองการคลิกเมาส์จริงที่ตัวปุ่ม
            btn_new.click_input()
            if log_fn: log_fn("Successfully clicked 'Add' button", color="#3ddc84")
            return True
        else:
            if log_fn: log_fn("Could not find 'Add' button", color="#ff4f4f")
            return False

    except Exception as e:
        if log_fn: log_fn(f"Error clicking Add button: {str(e)}", color="#ff4f4f")
        return False

def GetFirstRedErrorCode(main_form, log_fn=None):
    # 1. Get the App window's current position on the screen
    app_rect = main_form.rectangle()
    
    # 2. Calculate the total width of the window
    app_w = app_rect.right - app_rect.left
    
    # 3. Define the Bounding Box (L, T, R, B)
    # We start at the left edge and end 40% of the way across
    capture_bbox = (
        app_rect.left + 5,      # เริ่มจากขอบซ้ายของหน้าต่างจริงๆ
        app_rect.top + 100,     # หลบ Header ของโปรแกรม
        app_rect.left + 250,    # แคปเฉพาะแถบแคบๆ ทางซ้ายที่แสดง Error Code
        app_rect.bottom - 5
    )                  # [BOTTOM] Go to the bottom edge

    
    # 4. Grab the image (This is ONLY the left 40% of the app)
    img = PIL.ImageGrab.grab(bbox=capture_bbox)
    img.save("red_detect.png") 
    
    # 5. Scan the pixels in this "Left Side" image
    img_w, img_h = img.size
    pixels = img.load()
    found_pos_local = None

    # Scan bottom-to-top for the Red row
    for y in range(0, img_h, 3):
        for x in range(0, img_w, 5):
            r, g, b = pixels[x, y][:3]
            # Red detection (F561 color)
            if r > 200 and g < 60 and b < 60:
                found_pos_local = (x, y)
                break
        if found_pos_local: break

    # 6. Click the row if found
    if found_pos_local:
        screen_x = capture_bbox[0] + found_pos_local[0]
        screen_y = capture_bbox[1] + found_pos_local[1] #ship to click RED

        if log_fn: log_fn(f"Targeted RED row at {screen_x}, {screen_y}")
        time.sleep(0.3)  # small delay before click
        mouse.click(button='left', coords=(screen_x, screen_y))
        ClickAddRepairButton(main_form, log_fn)
        return (screen_x, screen_y)

    return None

def FillRepairForm(main_form, data, log_fn=None):
    print("--- [ DEBUG: DATA FROM EXCEL ] ---")
    if data:
        print(f"Full Data: {data}")
        print(f"Phenomenon Value: '{data['phenomenon']}'")
        print(f"First Letter: '{data['phenomenon'].strip()[:1].lower()}'")
        
        # เช็ค Shortcut ที่จะใช้
        char = data['phenomenon'].strip()[:1].lower()
        if char == 'a': print("Action: Will press Ctrl+A")
        elif char == 's': print("Action: Will press Ctrl+S")
        elif char == 'e': print("Action: Will press Ctrl+E")
        else: print(f"Action: Will Type '{data['phenomenon']}' directly")
    else:
        print("Error: No data returned from ReadData.GetFormData")
    print("----------------------------------")

    try:
        # 1. คลิกที่ช่อง Phenomenon เพื่อให้มั่นใจว่า Cursor อยู่ตรงนั้น
        # จากรูป ช่องแรกสุดมักจะเป็น TComboBox
        phenom_field = main_form.child_window(class_name="TComboBox", found_index=0)
        phenom_field.click_input() 
        time.sleep(0.5)

        # 2. ดึงข้อมูลจาก Excel (ดึงตัวอักษรแรกมาเช็ค)
        phenomenon_val = str(data.get("phenomenon", "")).strip()
        first_char = phenomenon_val[:1].lower()

        if log_fn: log_fn(f"Processing Phenomenon: {phenomenon_val}")

        # 3. ส่ง Shortcut ตามตัวอักษรแรก (a, s, e)
        if first_char == 'a':
            keyboard.send_keys("^a") # Ctrl + A
        elif first_char == 's':
            keyboard.send_keys("^s") # Ctrl + S
        elif first_char == 'e':
            keyboard.send_keys("^e") # Ctrl + E
        else:
            # ถ้าไม่มี Shortcut ให้พิมพ์ข้อความตรงๆ
            keyboard.write(phenomenon_val)
        
        time.sleep(0.3)
        keyboard.send_keys("{ENTER}") # ยืนยันการเลือก
        
        # 4. เลื่อนไปฟิลด์ถัดไป (Failure Code) โดยใช้ TAB
        keyboard.send_keys("{TAB 2}")
        time.sleep(0.2)
        keyboard.write(str(data.get("failure_code", "")))
        keyboard.send_keys("{TAB 3}")
        time.sleep(0.2)
        keyboard.write(str(data.get("location_code", "")))

        # 5. กรอกฟิลด์อื่นๆ ตามลำดับในหน้าจอ (ใช้ TAB ไล่ไปเรื่อยๆ)
        # Location Code, Reason Code, Handling ฯลฯ
        # ตัวอย่าง: กรอก Handling ในช่องด้านล่าง
        handling_val = data.get("handling", "")
        if handling_val:
            # ค้นหา TComboBox หรือ TEdit สำหรับ Handling (ดูจากรูปอยู่โซนล่างซ้าย)
            handling_field = main_form.child_window(title="Handling", control_type="Edit") # หรือใช้ index
            # หรือใช้การ TAB ไปเรื่อยๆ จนถึงช่องที่ต้องการ
            
        if log_fn: log_fn("Automation: Form filled via Shortcut Keys", color="#3ddc84")
        return True

    except Exception as e:
        if log_fn: log_fn(f"Fill Error: {str(e)}", color="#ff4f4f")
        return False



def RunRepairProcess(sn, log_fn, status_fn, form_data=None):
    try:
        status_fn("SCANNING", AMBER)
        log_fn(f"Serial Number: {sn}")

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

        app       = Application(backend="win32").connect(handle=window.handle)
        input_form = app.window(handle=window.handle)
        input_form.set_focus()

        sn_input = input_form.child_window(class_name="TEdit", found_index=0)
        sn_input.set_edit_text(sn)
        log_fn("Filled serial number")

        btn_repair = input_form.child_window(title="Repair", class_name="TBitBtn", found_index=0)
        btn_repair.click()
        log_fn("Clicked Repair button")

        status_fn("LOADING", AMBER)
        log_fn("Waiting for main form...")
        main_window = WaitForMainForm(timeout=10)

        repair_app  = Application(backend="win32").connect(handle=main_window.handle)
        repair_form = repair_app.window(handle=main_window.handle)
        repair_form.set_focus()
        time.sleep(1)
        log_fn("Main form ready")

        status_fn("DETECTING", AMBER)
        log_fn("Scanning for red error codes...")
        code = GetFirstRedErrorCode(repair_form)

        if ClickAddRepairButton(repair_form, log_fn):
            time.sleep(0.5) # รอให้ Form กรอกข้อมูลโหลดขึ้นมา
            # เรียกฟังก์ชันกรอกฟอร์ม
            FillRepairForm(repair_form, form_data, log_fn)

        if code:
            log_fn(f"Error code found: {code}", color=RED_ERR)
            status_fn("ERROR FOUND", RED_ERR)
        else:
            log_fn("No red error codes found", color=GREEN)
            status_fn("PASS", GREEN)

        return code

    except Exception as e:
        log_fn(f"[EXCEPTION] {e}", color=RED_ERR)
        status_fn("FAILED", RED_ERR)
        return None


def SelectPhenomenon(repair_win_handle, value):
    try:
        app = Desktop(backend="uia")
        form = app.window(handle=repair_win_handle.handle)
        form.set_focus()
        time.sleep(0.3)

        all_combos = []
        for cb in form.descendants():
            try:
                if cb.class_name() == "TComboBox" and cb.is_visible():
                    rect = cb.rectangle()
                    all_combos.append((cb, rect))
            except Exception:
                pass

        if not all_combos:
            print("[ERROR] No TComboBox found")
            return False

        print(f"[DEBUG] Found {len(all_combos)} combo(s)")
        for i, (_, rect) in enumerate(all_combos):
            print(f"  Combo[{i}] top={rect.top}, left={rect.left}")

        target_cb = min(all_combos, key=lambda item: (item[1].top, item[1].left))[0]
        print("[INFO] Selected top-left combo as Phenomenon")

        rect = target_cb.rectangle()
        target_cb.click_input(coords=(rect.width() // 2, rect.height() // 2))
        time.sleep(0.2)
        target_cb.set_focus()
        time.sleep(0.2)

        try:
            target_cb.select(value)
            print(f"[INFO] Selected '{value}' via select()")
            return True
        except Exception:
            print("[DEBUG] select() failed -> fallback")

        first_char = value[0] if value else ""
        if first_char:
            form.set_focus()
            time.sleep(0.2)
            print(f"[ACTION] Sending key '{first_char}'")
            form.type_keys(first_char)
            time.sleep(0.1)
            form.type_keys("{ENTER}")
            time.sleep(0.1)
            form.type_keys("{TAB}")
            return True

        return False

    except Exception as e:
        print(f"[ERROR] SelectPhenomenon failed: {e}")
        return False


def FillLocationCode(repair_win_handle, location_code):
    try:
        location_code = str(location_code).strip()

        if not location_code:
            print("[INFO] Location Code is empty, skipping")
            return True

        form = Desktop(backend="uia").window(handle=repair_win_handle.handle)
        form.set_focus()
        time.sleep(0.2)

        for _ in range(3):
            keyboard.send_keys("{TAB}")
            time.sleep(0.1)

        keyboard.send_keys(location_code, with_spaces=True)
        time.sleep(0.1)
        keyboard.send_keys("{TAB}")
        print(f"[INFO] Filled Location Code: '{location_code}'")
        return True

    except Exception as e:
        print(f"[ERROR] FillLocationCode failed: {e}")
        return False


def FillFailureCode(repair_win_handle, failure_code):
    try:
        failure_code = str(failure_code).strip()

        if not failure_code:
            print("[INFO] Failure Code is empty, skipping")
            return True

        form = Desktop(backend="uia").window(handle=repair_win_handle.handle)
        form.set_focus()
        time.sleep(0.2)

        keyboard.send_keys(failure_code, with_spaces=True)
        keyboard.send_keys("{TAB}")

        keyboard.send_keys(failure_code, with_spaces=True)
        time.sleep(0.1)
        keyboard.send_keys("{TAB}")
        print(f"[INFO] Filled Failure Code: '{failure_code}'")
        return True

    except Exception as e:
        print(f"[ERROR] FillFailureCode failed: {e}")
        return False


def ClickOK(form):
    try:
        print("[INFO] Clicked OK")
        time.sleep(0.3)
    except Exception as e:
        print(f"[ERROR] Cannot click OK: {e}")
        raise


def GetFirstRedErrorCode(main_form, form_data):
    SetDPIAware()

    code = None
    found_red_row = False
    phenomenon_value = ""
    failure_code = ""
    location_code = ""
    if form_data:
        phenomenon_value = str(form_data.get("phenomenon", "")).strip()
        failure_code = str(form_data.get("failure_code", "")).strip()
        location_code = str(form_data.get("location_code", "")).strip()

    target_grid = FindErrorCodeDBGrid(main_form)
    if not target_grid:
        print("[ERROR] Cannot find Error Code TDBGrid")
        return None, False

    grid_rect = target_grid.rectangle()
    img = GrabWindow(main_form.handle, rect=grid_rect)
    pixels = img.load()
    width, height = img.size

    first_red_x = None
    first_red_y = None
    for y in range(height):
        row_red_x = [x for x in range(width) if is_red_bg(*pixels[x, y])]
        red_count = len(row_red_x)
        if red_count > width * 0.1:
            first_red_x = row_red_x[0]
            first_red_y = y
            found_red_row = True
            break

    if not found_red_row:
        print("[INFO] No red highlighted rows found (Machine PASS)")
        return None, False

    screen_x = grid_rect.left + first_red_x + 20
    screen_y = grid_rect.top + first_red_y + 20 #ship from green to red
    print(f"[INFO] First red row detected at screen=({screen_x}, {screen_y})")

    main_rect = main_form.rectangle()
    rel_x = screen_x - main_rect.left
    rel_y = screen_y - main_rect.top

    print(f"[DEBUG] Clicking red row at screen=({screen_x}, {screen_y}) rel=({rel_x}, {rel_y})")

    try:
        main_form.click_input(coords=(rel_x, rel_y))
    except Exception:
        mouse.click(button="left", coords=(screen_x, screen_y))
    time.sleep(0.5)

    error_code_edit = FindErrorCodeEdit(main_form)
    if error_code_edit:
        raw_text = error_code_edit.texts()[0].strip()
        if raw_text:
            code = raw_text
            print(f"[INFO] Successfully read error code: '{code}'")
        else:
            print("[WARNING] Found TDBEdit but it is EMPTY")
    else:
        print("[ERROR] TDBEdit control not found")

    try:
        add_btn = main_form.child_window(title="Add", class_name="TBitBtn")
        if add_btn.exists():
            add_btn.click()
            print("[INFO] Clicked Add button")

            repair_win = WaitForRepairWindow(timeout=5)
            if repair_win:
                SelectPhenomenon(repair_win, phenomenon_value)
                FillFailureCode(repair_win, failure_code)
                FillLocationCode(repair_win, location_code)
                rapp = Application(backend="uia").connect(handle=repair_win.handle)
                rform = rapp.window(handle=repair_win.handle)
                ClickOK(rform)
            else:
                print("[ERROR] Cannot find Repair Window")
    except Exception as e:
        print(f"[DEBUG] Add/Repair flow skipped or error: {e}")

    return code, found_red_row


def RunRepairProcess(sn, log_fn, status_fn, from_data=None):
    try:
        form_data = from_data or ReadData.GetFormData(sn)

        status_fn("SCANNING", AMBER)
        log_fn(f"Serial Number: {sn}")

        repair_windows = Desktop(backend="uia").windows(
            title_re=r"^Repair-Rev",
            top_level_only=True,
            visible_only=True,
        )
        if not repair_windows:
            raise RuntimeError("Repair-Rev window not found")

        foreground_handle = ctypes.windll.user32.GetForegroundWindow()
        window = next(
            (w for w in repair_windows if w.handle == foreground_handle),
            repair_windows[0],
        )

        app = Application(backend="uia").connect(handle=window.handle)
        input_form = app.window(handle=window.handle)
        input_form.set_focus()

        sn_input = input_form.child_window(class_name="TEdit", found_index=0)
        sn_input.set_edit_text(sn)
        log_fn("Filled serial number")

        input_form.child_window(
            title="Repair",
            class_name="TBitBtn",
            found_index=0,
        ).click()
        log_fn("Clicked Repair button")

        status_fn("LOADING", AMBER)
        log_fn("Waiting for main form...")
        main_window = WaitForMainForm(timeout=10)

        repair_app = Application(backend="uia").connect(handle=main_window.handle)
        repair_form = repair_app.window(handle=main_window.handle)
        repair_form.set_focus()
        time.sleep(1)
        log_fn("Main form ready")

        status_fn("DETECTING", AMBER)
        log_fn("Scanning for red error codes...")
        code, red_row_exists = GetFirstRedErrorCode(repair_form, form_data)

        if red_row_exists:
            if code:
                log_fn(f"Error code found: {code}", color=RED_ERR)
                status_fn("ERROR FOUND", RED_ERR)
            else:
                log_fn("Detected red row but failed to read code text", color=RED_ERR)
                status_fn("READ ERROR", RED_ERR)
        else:
            log_fn("No red error codes found", color=GREEN)
            status_fn("PASS", GREEN)

        return code

    except Exception as e:
        log_fn(f"[EXCEPTION] {e}", color=RED_ERR)
        status_fn("FAILED", RED_ERR)
        return None

def GetRepairWindowRect():
    """Get the bounding rect of the Repair Window."""
    repair_win = Desktop(backend="win32").window(title="Repair Window")
    repair_win.wait("visible", timeout=10)
    repair_win.set_focus()
    time.sleep(0.3)
    r = repair_win.rectangle()
    return repair_win, r

def CaptureWindow(rect):
    """Screenshot just the Repair Window."""
    img = PIL.ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
    return img

def ClickInputNextToLabel(win_rect, label_box, img, log_fn=print):
    """
    Given a label bounding box (relative to window screenshot),
    click the INPUT FIELD that sits to its right on the same row.
    
    Strategy: scan the image pixels to the right of the label
    looking for a white/light-yellow rectangle (input field).
    """
    lx, ly, lw, lh = label_box
    img_w, img_h = img.size
    pixels = img.load()

    # Search rightward from end of label, on the same vertical band
    search_y_center = ly + lh // 2
    search_y_range  = range(max(0, search_y_center - 4), min(img_h, search_y_center + 4))

    input_x = None
    for x in range(lx + lw + 5, min(img_w, lx + lw + 300)):
        for y in search_y_range:
            r, g, b = pixels[x, y][:3]
            # Input fields are white or light yellow (Delphi default)
            if r > 220 and g > 220 and b > 180:
                input_x = x
                break
        if input_x:
            break

    if input_x is None:
        log_fn(f"  Could not find input field to the right of label")
        return False

    # Convert image-relative coords → screen coords
    screen_x = win_rect.left + input_x + 10   # +10 to land inside the field
    screen_y = win_rect.top

    from pywinauto import mouse
    mouse.click(button="left", coords=(screen_x, screen_y))
    time.sleep(0.15)
    return True

# ============================================================
#  GUI
# ============================================================

class RepairGUI:
    def __init__(self, root):
        self.root  = root
        self.root.title("Repair Automation")
        self.root.configure(bg=BG_DARK)
        self.root.resizable(False, False)
        self.root.geometry("620x720")

        self._build_fonts()
        self._build_ui()
        self._set_ready()

        # focus barcode input เสมอ
        self.sn_entry.focus_set()

    # ── fonts ───────────────────────────────────────────────
    def _build_fonts(self):
        self.f_title  = tkfont.Font(family="Consolas", size=13, weight="bold")
        self.f_label  = tkfont.Font(family="Consolas", size=9)
        self.f_entry  = tkfont.Font(family="Consolas", size=16, weight="bold")
        self.f_status = tkfont.Font(family="Consolas", size=22, weight="bold")
        self.f_log    = tkfont.Font(family="Consolas", size=9)
        self.f_btn    = tkfont.Font(family="Consolas", size=10, weight="bold")
        self.f_badge  = tkfont.Font(family="Consolas", size=8)

    # ── UI builder ──────────────────────────────────────────
    def _build_ui(self):
        pad = dict(padx=20)

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
        status_frame = tk.Frame(self.root, bg=BG_PANEL,
                                highlightbackground=BORDER, highlightthickness=1)
        status_frame.pack(fill="x", padx=20, pady=(16, 0))

        inner = tk.Frame(status_frame, bg=BG_PANEL)
        inner.pack(pady=18)

        tk.Label(inner, text="STATUS", font=self.f_badge,
                 bg=BG_PANEL, fg=TEXT_SEC).pack()

        self.status_lbl = tk.Label(inner, text="READY", font=self.f_status,
                                   bg=BG_PANEL, fg=GREEN)
        self.status_lbl.pack()

        # ── SN input ─────────────────────────────────────────
        tk.Label(self.root, text="SERIAL NUMBER", font=self.f_badge,
                 bg=BG_DARK, fg=TEXT_SEC, anchor="w").pack(fill="x", padx=20, pady=(16, 2))

        entry_frame = tk.Frame(self.root, bg=AMBER, padx=2, pady=2)
        entry_frame.pack(fill="x", padx=20)

        self.sn_var   = tk.StringVar()
        self.sn_entry = tk.Entry(
            entry_frame,
            textvariable=self.sn_var,
            font=self.f_entry,
            bg=BG_INPUT, fg=TEXT_PRI,
            insertbackground=AMBER,
            relief="flat",
            bd=0
        )
        self.sn_entry.pack(fill="x", ipady=10, ipadx=12)
        self.sn_entry.bind("<Return>", self._on_scan)
        self.sn_entry.bind("<FocusOut>", lambda e: self.sn_entry.focus_set())

        tk.Label(self.root, text="Scan barcode or type SN then press Enter",
                 font=self.f_badge, bg=BG_DARK, fg=TEXT_SEC).pack(anchor="w", padx=20, pady=(4, 0))

        # ── result row ───────────────────────────────────────
        res_frame = tk.Frame(self.root, bg=BG_DARK)
        res_frame.pack(fill="x", padx=20, pady=(14, 0))

        tk.Label(res_frame, text="ERROR CODE", font=self.f_badge,
                 bg=BG_DARK, fg=TEXT_SEC).pack(anchor="w")

        self.result_lbl = tk.Label(res_frame, text="—", font=self.f_entry,
                                   bg=BG_DARK, fg=TEXT_SEC, anchor="w")
        self.result_lbl.pack(anchor="w")

        # ── buttons ──────────────────────────────────────────
        btn_row = tk.Frame(self.root, bg=BG_DARK)
        btn_row.pack(fill="x", padx=20, pady=(14, 0))

        self.reset_btn = tk.Button(
            btn_row,
            text="↺  RESET",
            font=self.f_btn,
            bg=BG_PANEL, fg=AMBER,
            activebackground=AMBER_DIM, activeforeground=TEXT_PRI,
            relief="flat", bd=0, cursor="hand2",
            command=self._reset,
            padx=18, pady=8
        )
        self.reset_btn.pack(side="left")

        self.run_btn = tk.Button(
            btn_row,
            text="▶  RUN",
            font=self.f_btn,
            bg=AMBER, fg=BG_DARK,
            activebackground=AMBER_DIM, activeforeground=TEXT_PRI,
            relief="flat", bd=0, cursor="hand2",
            command=lambda: self._on_scan(None),
            padx=18, pady=8
        )
        self.run_btn.pack(side="right")

        # ── log panel ────────────────────────────────────────
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x", padx=20, pady=(16, 0))

        log_header = tk.Frame(self.root, bg=BG_DARK)
        log_header.pack(fill="x", padx=20, pady=(8, 0))

        tk.Label(log_header, text="PROCESS LOG", font=self.f_badge,
                 bg=BG_DARK, fg=TEXT_SEC).pack(side="left")

        tk.Button(log_header, text="CLEAR", font=self.f_badge,
                  bg=BG_DARK, fg=TEXT_SEC, relief="flat", bd=0,
                  cursor="hand2", command=self._clear_log).pack(side="right")

        log_frame = tk.Frame(self.root, bg=BG_PANEL,
                             highlightbackground=BORDER, highlightthickness=1)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(4, 20))

        self.log_text = tk.Text(
            log_frame,
            font=self.f_log,
            bg=BG_PANEL, fg=TEXT_MONO,
            insertbackground=AMBER,
            relief="flat", bd=0,
            state="disabled",
            wrap="word",
            padx=10, pady=8
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scroll = tk.Scrollbar(log_frame, command=self.log_text.yview,
                              bg=BG_PANEL, troughcolor=BG_PANEL,
                              activebackground=BORDER, relief="flat", bd=0)
        scroll.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scroll.set)

        # tag colors สำหรับ log
        self.log_text.tag_config("green",  foreground=GREEN)
        self.log_text.tag_config("red",    foreground=RED_ERR)
        self.log_text.tag_config("amber",  foreground=AMBER)
        self.log_text.tag_config("dim",    foreground=TEXT_SEC)
        self.log_text.tag_config("normal", foreground=TEXT_MONO)

    # ── clock ────────────────────────────────────────────────
    def _tick_clock(self):
        self.clock_lbl.config(text=datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))
        self.root.after(1000, self._tick_clock)

    # ── state helpers ────────────────────────────────────────
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

    # ── log helpers ──────────────────────────────────────────
    def _log(self, message, color=None):
        """thread-safe log append"""
        def _write():
            ts  = datetime.now().strftime("%H:%M:%S")
            tag = {GREEN: "green", RED_ERR: "red", AMBER: "amber"}.get(color, "normal")

            self.log_text.config(state="normal")
            self.log_text.insert("end", f"[{ts}]  ", "dim")
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

    # ── actions ──────────────────────────────────────────────
    def _on_scan(self, event):
        if self.running:
            return
        sn = self.sn_var.get().strip()
        if not sn:
            self._log("No serial number entered", color=RED_ERR)
            return

        self._set_busy()
        self._set_status("RUNNING...", AMBER)
        self._set_result("—", TEXT_SEC)
        self._log(f"── START ─────────────────────", color=AMBER)

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
        """callback after process — clear SN and waiting next scan"""
        self.sn_var.set("")
        self._set_ready()
        self.sn_entry.focus_set()
        self._log("── DONE — ready for next scan ─", color=AMBER)

    def _reset(self):
        """reset to initial state"""
        if self.running:
            return
        self.sn_var.set("")
        self._clear_log()
        self._set_status("READY", GREEN)
        self._set_result("—", TEXT_SEC)
        self.sn_entry.focus_set()
        self._log("System reset", color=AMBER)


# ============================================================
#  ENTRY POINT
# ============================================================

if __name__ == "__main__":
    CheckPythonBitness()
    root = tk.Tk()
    app  = RepairGUI(root)
    root.mainloop()
