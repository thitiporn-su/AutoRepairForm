import ctypes
import sys

# ── ต้องเรียกก่อน import อื่นทั้งหมด ────────────────────────
def SetDPIAware():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

SetDPIAware()

# ── imports ──────────────────────────────────────────────────
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

try:
    import ReadData
except ImportError:
    ReadData = None


# ============================================================
#  CONFIG — ตั้งค่าตรงนี้
# ============================================================
PHENOMENON_VALUE = "Appearance"   # ค่าที่ต้องการเลือกใน Phenomenon dropdown


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
#  GRAB WINDOW (PrintWindow API — ไม่ขึ้นกับ monitor)
# ============================================================
def GrabWindow(hwnd, rect=None):
    """
    Capture จาก window โดยตรงผ่าน PrintWindow API
    - ไม่ขึ้นกับ monitor position / multi-monitor
    - ได้ภาพแม้ window ถูก overlap บัง
    hwnd : handle ของ main_form
    rect : pywinauto rectangle — crop เฉพาะส่วนนี้
    """
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    win_w = right  - left
    win_h = bottom - top

    hwnd_dc = win32gui.GetWindowDC(hwnd)
    mfc_dc  = win32ui.CreateDCFromHandle(hwnd_dc)
    save_dc = mfc_dc.CreateCompatibleDC()

    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc_dc, win_w, win_h)
    save_dc.SelectObject(bmp)

    # PW_RENDERFULLCONTENT = 2 (รองรับ layered/DX window)
    result = ctypes.windll.user32.PrintWindow(hwnd, save_dc.GetSafeHdc(), 2)
    print(f"[DEBUG] PrintWindow result={result} (1=success)")

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
        crop_l = rect.left   - left
        crop_t = rect.top    - top
        crop_r = rect.right  - left
        crop_b = rect.bottom - top
        img = img.crop((crop_l, crop_t, crop_r, crop_b))
        print(f"[DEBUG] Cropped {img.width}x{img.height}px "
              f"(rel: {crop_l},{crop_t},{crop_r},{crop_b})")

    img.save("debug_grid.png")
    print(f"[DEBUG] Saved debug_grid.png ({img.width}x{img.height}px)")
    return img


# ============================================================
#  COLOR HELPER
# ============================================================
def is_red_bg(r, g, b):
    """ตรวจสอบ pixel สีแดง background — tolerance กว้างรองรับทุกจอ"""
    return (
        r > 150 and
        g < 100 and
        b < 100 and
        r > g * 2 and
        r > b * 2
    )


# ============================================================
#  WINDOW HELPERS
# ============================================================
def WaitForMainForm(timeout=10):
    """รอ TfrmMain หลัง click Repair"""
    print("[INFO] Waiting for TfrmMain...")
    start = time.time()
    while time.time() - start < timeout:
        for w in Desktop(backend="win32").windows(title_re=r"^Repair-Rev"):
            if w.class_name() == "TfrmMain":
                print("[INFO] TfrmMain found!")
                return w
        time.sleep(0.5)
    raise RuntimeError(f"Timeout {timeout}s: TfrmMain did not appear")


def WaitForRepairWindow(timeout=10):
    """รอ Repair Window popup หลัง click Add"""
    print("[INFO] Waiting for Repair Window...")
    start = time.time()
    while time.time() - start < timeout:
        wins = Desktop(backend="win32").windows(title_re=r"^Repair Window")
        if wins:
            print("[INFO] Repair Window found!")
            return wins[0]
        time.sleep(0.5)
    raise RuntimeError(f"Timeout {timeout}s: Repair Window did not appear")


# ============================================================
#  FIND CONTROLS
# ============================================================
def FindErrorCodeDBGrid(main_form):
    """
    หา TDBGrid ที่เป็น Error Code List แบบ dynamic
    Strategy 1: หา grid ที่อยู่เหนือ TBitBtn New/Remove
    Strategy 2: fallback grid ซ้ายสุด
    """
    all_grids = []
    for child in main_form.descendants():
        try:
            if child.class_name() == "TDBGrid":
                all_grids.append(child)
        except Exception:
            pass

    print(f"[DEBUG] Found {len(all_grids)} TDBGrid(s)")
    for i, g in enumerate(all_grids):
        print(f"  Grid[{i}] rect={g.rectangle()}")

    if not all_grids:
        return None

    # Strategy 1: anchor จาก New/Remove button
    for child in main_form.descendants():
        try:
            if child.class_name() == "TBitBtn" and child.texts()[0] in ("New", "Remove"):
                btn_rect = child.rectangle()
                for grid in all_grids:
                    r = grid.rectangle()
                    if r.bottom <= btn_rect.top and abs(r.left - btn_rect.left) < 50:
                        print(f"[INFO] Found Error Code Grid (anchor): {r}")
                        return grid
        except Exception:
            pass

    # Strategy 2: grid ซ้ายสุด
    leftmost = min(all_grids, key=lambda g: g.rectangle().left)
    print(f"[INFO] Found Error Code Grid (fallback): {leftmost.rectangle()}")
    return leftmost


def FindErrorCodeEdit(main_form):
    all_edits = []
    for child in main_form.descendants():
        try:
            # กรองเอาเฉพาะ TDBEdit ที่มองเห็นได้
            if child.class_name() == "TDBEdit" and child.is_visible():
                all_edits.append(child)
        except Exception:
            pass

    print(f"[DEBUG] Found {len(all_edits)} visible TDBEdit(s)")

    # Strategy: ค้นหาตัวที่มีลักษณะเหมือน Error Code (เช่น F561)
    # เราจะข้ามตัวที่มีช่องว่าง หรือตัวที่เป็นประโยคยาวๆ (Description)
    for edit in all_edits:
        try:
            text = edit.texts()[0].strip()
            if not text:
                continue
                
            # ตรวจสอบ Pattern: 
            # 1. ถ้าขึ้นต้นด้วย 'F' (เช่น F561)
            # 2. หรือถ้าความยาวตัวอักษรไม่มากเกินไป (Error Code มักจะไม่ยาวเท่า Description)
            if text.startswith('F') or len(text) <= 12:
                # ลองเช็คว่ามันใช่ Description หรือเปล่า (ถ้ามีคำว่า 'height' หรือ 'ball' ให้ข้าม)
                if "height" in text.lower() or " " in text: 
                    continue
                    
                print(f"[INFO] Identified Error Code Edit: '{text}'")
                return edit
        except Exception:
            pass

    # Fallback: ถ้าหาตาม Pattern ไม่เจอจริงๆ ให้ลองเลี่ยง Index ที่มักจะเป็น Description
    # จากรูป Description อยู่ถัดจาก Error Code ดังนั้น Error Code น่าจะเป็น index 0 หรือ 1
    if len(all_edits) > 0:
        return all_edits[0] 

    return None

# ============================================================
#  REPAIR WINDOW ACTIONS
# ============================================================
# def SelectPhenomenon(repair_win_handle, value):
#     try:
#         from pywinauto import Desktop
#         app = Desktop(backend="win32")
#         form = app.window(handle=repair_win_handle.handle)
#         form.set_focus()

#         # 1. หาตัวที่อยู่บนสุด (Phenomenon)
#         all_combos = form.descendants(class_name="TComboBox")
#         target_cb = min(all_combos, key=lambda cb: cb.rectangle().top)

#         # 2. ดึงตัวอักษรตัวแรกของค่าที่จะใส่ (เช่น 'B' จาก 'Ball height')
#         first_char = value[0] if value else ""
        
#         if target_cb and first_char:
#             print(f"[ACTION] Sending '{first_char}' to select value starting with it.")
            
#             # คลิกเพื่อ Focus
#             target_cb.click_input()
            
#             # ส่งตัวอักษรตัวแรก -> กด Enter (เพื่อเลือก) -> กด Tab (เพื่อข้ามไปช่องถัดไป)
#             # {ENTER} ยืนยันรายการที่มันเด้งไปหา, {TAB} เลื่อนไปช่อง Failure Code
#             target_cb.type_keys(first_char + "{ENTER}{TAB}", pause=0.1)
            
#             return True
#         return False

#     except Exception as e:
#         print(f"[DEBUG] Error: {e}")
#         return False

def SelectPhenomenon(repair_win_handle, value):
    try:
        from pywinauto import Desktop
        import time

        app = Desktop(backend="win32")
        form = app.window(handle=repair_win_handle.handle)
        form.set_focus()
        time.sleep(0.3)

        # =========================================================
        # STEP 1: หา TComboBox ทั้งหมด
        # =========================================================
        all_combos = []

        for cb in form.descendants():
            try:
                if cb.class_name() == "TComboBox" and cb.is_visible():
                    r = cb.rectangle()
                    all_combos.append((cb, r))
            except Exception:
                pass

        if not all_combos:
            print("[ERROR] No TComboBox found")
            return False

        print(f"[DEBUG] Found {len(all_combos)} combo(s)")

        for i, (_, r) in enumerate(all_combos):
            print(f"  Combo[{i}] top={r.top}, left={r.left}")

        # =========================================================
        # STEP 2: เลือก combo "ซ้ายสุด + บนสุด"
        # =========================================================
        target_cb = min(all_combos, key=lambda x: (x[1].top, x[1].left))[0]

        print("[INFO] Selected top-left combo as Phenomenon")

        # =========================================================
        # STEP 3: click + focus
        # =========================================================
        r = target_cb.rectangle()

        target_cb.click_input(coords=(
            r.width() // 2,
            r.height() // 2
        ))
        time.sleep(0.2)

        target_cb.set_focus()
        time.sleep(0.2)

        # =========================================================
        # STEP 4: ลอง select ตรง
        # =========================================================
        try:
            target_cb.select(value)
            print(f"[INFO] Selected '{value}' via select()")
            return True
        except Exception:
            print("[DEBUG] select() failed → fallback")

        # =========================================================
        # STEP 5: ยิง key ผ่าน form (เสถียรสุด)
        # =========================================================
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

def ClickOK(form):
    try:
        # form.child_window(title="OK", class_name="TBitBtn").click()
        print("[INFO] Clicked OK")
        time.sleep(0.3)
    except Exception as e:
        print(f"[ERROR] Cannot click OK: {e}")
        raise


# ============================================================
#  CORE: FIND RED ERROR + CLICK ADD + REPAIR WINDOW
# ============================================================
def GetFirstRedErrorCode(main_form, phenomenon_value):
    """
    1. Capture grid จาก window โดยตรง (PrintWindow)
    2. Scan pixel หา red row แรก
    3. Return (code, found_red_row) เพื่อแยกแยะระหว่าง 'เครื่องปกติ' กับ 'อ่านค่าไม่ได้'
    """
    SetDPIAware()
    
    # --- เริ่มต้นสถานะ ---
    code = None
    found_red_row = False

    # ── หา TDBGrid ──────────────────────────────────────────
    target_grid = FindErrorCodeDBGrid(main_form)
    if not target_grid:
        print("[ERROR] Cannot find Error Code TDBGrid")
        return None, False

    grid_rect = target_grid.rectangle()

    # ── Capture จาก window โดยตรง ───────────────────────────
    img = GrabWindow(main_form.handle, rect=grid_rect)
    pixels = img.load()
    width, height = img.size

    # ── Scan หา red row แรก ──────────────────────────────────
    first_red_y = None
    for y in range(height):
        red_count = sum(1 for x in range(width) if is_red_bg(*pixels[x, y]))
        # ถ้าเจอสีแดงมากกว่า 10% ของความกว้างแถว
        if red_count > width * 0.1:
            first_red_y = grid_rect.top + y
            found_red_row = True
            break

    if not found_red_row:
        print("[INFO] No red highlighted rows found (Machine PASS)")
        return None, False

    print(f"[INFO] First red row detected at screen Y={first_red_y}")

    # ── Click เลือกแถวสีแดง ────────────────────────────────────
    main_rect = main_form.rectangle()
    click_x = grid_rect.left + (grid_rect.right - grid_rect.left) // 2
    
    # คลิกเพื่อให้ UI อัปเดตค่าเข้า TDBEdit
    main_form.click_input(coords=(
        click_x - main_rect.left,
        first_red_y - main_rect.top,
    ))
    time.sleep(0.5) # รอให้ Text ใน Edit เปลี่ยนค่า

    # ── อ่าน Error Code (ปรับปรุงให้หาตัวที่มี Text) ──────────────
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

    # ── Click Add ────────────────────────────────────────────
    try:
        add_btn = main_form.child_window(title="Add", class_name="TBitBtn")
        if add_btn.exists():
            add_btn.click()
            print("[INFO] Clicked Add button")
            
            # ── Process Repair Window ──
            repair_win = WaitForRepairWindow(timeout=5)
            if repair_win:
                SelectPhenomenon(repair_win, phenomenon_value)
                
                # หาปุ่ม OK ใน Repair Window
                rapp = Application(backend="win32").connect(handle=repair_win.handle)
                rform = rapp.window(handle=repair_win.handle)
                ClickOK(rform)
            else:
                print("[ERROR] Cannot find Repair Window")

    except Exception as e:
        print(f"[DEBUG] Add/Repair flow skipped or error: {e}")

    return code, found_red_row

# ============================================================
#  MAIN PROCESS (called from GUI thread)
# ============================================================
def RunRepairProcess(sn, log_fn, status_fn):
    try:
        status_fn("SCANNING", AMBER)
        log_fn(f"Serial Number: {sn}")

        # ── หา TFrmInputSN ───────────────────────────────────
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

        app        = Application(backend="win32").connect(handle=window.handle)
        input_form = app.window(handle=window.handle)
        input_form.set_focus()

        # ── กรอก SN ──────────────────────────────────────────
        sn_input = input_form.child_window(class_name="TEdit", found_index=0)
        sn_input.set_edit_text(sn)
        log_fn("Filled serial number")

        # ── Click Repair ─────────────────────────────────────
        input_form.child_window(
            title="Repair", class_name="TBitBtn", found_index=0
        ).click()
        log_fn("Clicked Repair button")

        # ── รอ TfrmMain ──────────────────────────────────────
        status_fn("LOADING", AMBER)
        log_fn("Waiting for main form...")
        main_window = WaitForMainForm(timeout=10)

        repair_app  = Application(backend="win32").connect(handle=main_window.handle)
        repair_form = repair_app.window(handle=main_window.handle)
        repair_form.set_focus()
        time.sleep(1)
        log_fn("Main form ready")

        # ── หา red error code ────────────────────────────────
        status_fn("DETECTING", AMBER)
        log_fn("Scanning for red error codes...")

        code, red_row_exists =GetFirstRedErrorCode(repair_form, PHENOMENON_VALUE)

        if red_row_exists:
            if code:
                log_fn(f"Error code found: {code}", color=RED_ERR)
                status_fn("ERROR FOUND", RED_ERR)
            else:
                # กรณีเจอแถวแดงแต่ Code เป็น None/ว่าง (เช่น เคส F156 ที่คุณเจอ)
                log_fn("Detected Red Row but FAILED to read Code text", color=RED_ERR)
                status_fn("READ ERROR", RED_ERR)
        else:
            # กรณีไม่เจอแถวแดงเลยจริงๆ
            log_fn("No red error codes found", color=GREEN)
            status_fn("PASS", GREEN)

        return code

    except Exception as e:
        log_fn(f"[EXCEPTION] {e}", color=RED_ERR)
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
        self.root.geometry("620x720")

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

        tk.Frame(self.root, bg=BORDER, height=1).pack(
            fill="x", padx=20, pady=(10, 0))

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
        tk.Frame(self.root, bg=BORDER, height=1).pack(
            fill="x", padx=20, pady=(16, 0))
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
        self.log_text.tag_config("green",  foreground=GREEN)
        self.log_text.tag_config("red",    foreground=RED_ERR)
        self.log_text.tag_config("amber",  foreground=AMBER)
        self.log_text.tag_config("dim",    foreground=TEXT_SEC)
        self.log_text.tag_config("normal", foreground=TEXT_MONO)

    def _tick_clock(self):
        self.clock_lbl.config(
            text=datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))
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
            tag = {GREEN: "green", RED_ERR: "red",
                   AMBER: "amber"}.get(color, "normal")
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
        self.root.after(0, lambda: self.status_lbl.config(
            text=text, fg=color))

    def _set_result(self, text, color=TEXT_SEC):
        self.root.after(0, lambda: self.result_lbl.config(
            text=text, fg=color))

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
        self._log("── START ─────────────────────", color=AMBER)

        def worker():
            code = RunRepairProcess(
                sn=sn,
                log_fn=self._log,
                status_fn=self._set_status,
            )
            if code:
                self.root.after(0, lambda: self._set_result(code, RED_ERR))
            else:
                self.root.after(
                    0, lambda: self._set_result("No error found", GREEN))
            self.root.after(0, self._after_process)

        threading.Thread(target=worker, daemon=True).start()

    def _after_process(self):
        self.sn_var.set("")
        self._set_ready()
        self.sn_entry.focus_set()
        self._log("── DONE — ready for next scan ─", color=AMBER)

    def _reset(self):
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
    root = tk.Tk()
    app  = RepairGUI(root)
    root.mainloop()