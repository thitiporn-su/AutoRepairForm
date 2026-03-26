import tkinter as tk
from tkinter import font as tkfont
import threading
import time
import ctypes
import PIL.ImageGrab

from PIL import ImageGrab
from datetime import datetime
from pywinauto import Application, Desktop, mouse
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)
try:
    import ReadData
except ImportError:
    ReadData = None


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

def GetFirstRedErrorCode(main_form):
    # 1. Find the specific "Error Code List" grid (TDBGrid)
    target_grid = FindErrorCodeDBGrid(main_form)
    if not target_grid: 
        return None

    # 2. Get the screen coordinates of just this grid
    rect = target_grid.rectangle()
    
    # 3. Use PIL.ImageGrab on only this Bounding Box (bbox)
    # Note: We add a small margin to avoid the grid borders
    img = PIL.ImageGrab.grab(bbox=(
        rect.left + 5, 
        rect.top,  # Skip the "Error Code" header
        rect.right - 20, # Skip scrollbar
        rect.bottom - 5
    ))
    img.save("debug_grid_only.png") # Check this; it should be just the list

    width, height = img.size
    pixels = img.load()

    # 4. Scan the small image for red background
    for y in range(0, height, 5):
        for x in range(0, width, 5):
            r, g, b = pixels[x, y][:3]
            
            # Bright Red Detection (Matches F561 row in your image)
            if r > 200 and g < 60 and b < 60:
                # Calculate screen position to move mouse and click
                screen_x = rect.left + 5 + x
                screen_y = rect.top + 30 + y
                
                # Step A: Move and Click the red row
                mouse.move(coords=(screen_x, screen_y))
                time.sleep(0.1)
                mouse.click(button='left', coords=(screen_x, screen_y))
                
                # Step B: Click the "Add" button
                time.sleep(0.5)
                try:
                    btn_add = main_form.child_window(title="Add", class_name="TBitBtn")
                    btn_add.click()
                except:
                    pass
                
                # Step C: Extract the dynamic code (F561) from the Edit field
                for edit in main_form.descendants(class_name="TDBEdit"):
                    val = edit.window_text().strip()
                    if val and val != "N/A" and not (val.isdigit() and len(val) == 12):
                        return val
                return "RED_FOUND"

    return None



def RunRepairProcess(sn, log_fn, status_fn):
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