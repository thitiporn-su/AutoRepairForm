import time
import ctypes
import ReadData
import PIL.ImageGrab
from pywinauto import Application, Desktop
import win32con
import win32gui

def RedErrorExists(main_form):
    rect = main_form.rectangle()
    img = PIL.ImageGrab.grab(bbox=(rect.left, rect.top, rect.right, rect.bottom))
    width, height = img.size
    print(f"Scanning area {width}X{height}")

    def is_red(r, g, b):
        return r == 255 and g == 0 and b == 0

    for y in range(50, height): 
        for x in range(0, width-50):
            r, g, b = img.getpixel((x, y))
            
            if is_red(r, g, b):
                print(f"[FOUND RED] at ({x},{y}) -> RGB({r},{g},{b})")
                print(f"[CLICK] Red error detected at screen ({x},{y})")
                main_form.click_input(coords=(x, y))
                return True

            # window.click_input(coords=(x, y))
        
    return False

def DebugChildren(main_form):
    print("\n=== All children of main_form ===")
    for child in main_form.descendants():
        try:
            rect = child.rectangle()
            texts = child.texts()
            print(f"  class='{child.class_name()}' | text={texts} | rect={rect}")
        except:
            pass

def DebugAllWindows(main_form):
    """ดู window ทั้งหมดที่เปิดอยู่ รวม popup"""
    print("\n=== All top-level Repair windows ===")
    all_windows = Desktop(backend="win32").windows(title_re=r"^Repair-Rev")
    for w in all_windows:
        print(f"\n>> handle={w.handle} title='{w.window_text()}' class='{w.class_name()}'")
        print(f"   rect={w.rectangle()}")
        for child in w.descendants():
            try:
                print(f"     child class='{child.class_name()}' text={child.texts()[:2]} rect={child.rectangle()}")
            except:
                pass

def WaitForMainForm(timeout=10):
    """รอจนกว่า TfrmMain จะปรากฏ"""
    start = time.time()
    while time.time() - start < timeout:
        all_windows = Desktop(backend="win32").windows(title_re=r"^Repair-Rev")
        for w in all_windows:
            if w.class_name() == "TfrmMain":
                print("[INFO] TfrmMain found!")
                return w
        time.sleep(0.5)
    raise RuntimeError("Timeout: TfrmMain did not appear")

def GetRedErrorCodesFromDBGrid(main_form):
    # หา TDBGrid ฝั่งซ้าย (Error Code List)
    target_grid = None
    for child in main_form.descendants():
        try:
            if child.class_name() == "TDBGrid":
                r = child.rectangle()
                if r.left < 900 and r.top < 400:
                    target_grid = child
                    break
        except:
            pass

    if not target_grid:
        print("[ERROR] Cannot find Error Code TDBGrid")
        return []

    grid_rect = target_grid.rectangle()

    # Screenshot เฉพาะ TDBGrid
    img = PIL.ImageGrab.grab(bbox=(
        grid_rect.left, grid_rect.top,
        grid_rect.right, grid_rect.bottom
    ))
    pixels = img.load()
    width, height = img.size

    def is_red_bg(r, g, b):
        return r > 180 and g < 80 and b < 80

    # หา Y ของแต่ละ red row
    red_row_screen_ys = []
    in_red_band = False
    for y in range(height):
        red_count = sum(1 for x in range(width) if is_red_bg(*pixels[x, y]))
        if red_count > width * 0.2:
            if not in_red_band:
                red_row_screen_ys.append(grid_rect.top + y + 2)
                in_red_band = True
        else:
            in_red_band = False

    if not red_row_screen_ys:
        print("[INFO] No red rows found")
        return []

    # หา TDBEdit Error Code field
    error_code_edit = None
    for child in main_form.descendants():
        try:
            if child.class_name() == "TDBEdit":
                r = child.rectangle()
                if 1100 < r.left < 1200 and 250 < r.top < 300:
                    error_code_edit = child
                    break
        except:
            pass

    if not error_code_edit:
        print("[ERROR] Cannot find Error Code TDBEdit")
        return []

    # Click แต่ละ red row แล้วอ่าน Error Code
    results = []
    main_rect = main_form.rectangle()
    click_x = grid_rect.left + (grid_rect.right - grid_rect.left) // 2

    for screen_y in red_row_screen_ys:
        main_form.click_input(coords=(
            click_x - main_rect.left,
            screen_y - main_rect.top
        ))
        time.sleep(0.3)

        try:
            code = error_code_edit.texts()[0].strip()
            if code and code not in results:
                print(f"  → Error Code: '{code}'")
                results.append(code)
        except Exception as e:
            print(f"  [ERROR] {e}")

    return results

def main():
    DATA = ReadData.GetFormData()
    SN = DATA["serial_number"]

    time.sleep(3)

    repair_windows = Desktop(backend="win32").windows(
        title_re=r"^Repair-Rev",
        top_level_only=True,
        visible_only=True,
    )

    if not repair_windows:
        raise RuntimeError("Could not find a visible top-level 'Repair-Rev' window.")

    foreground_handle = ctypes.windll.user32.GetForegroundWindow()
    window = next((win for win in repair_windows if win.handle == foreground_handle), repair_windows[0])

    app = Application(backend="win32").connect(handle=window.handle)
    main_form = app.window(handle=window.handle)
    main_form.set_focus()

    SN_input = main_form.child_window(class_name="TEdit", found_index=0)
    SN_input.set_edit_text(SN)
    print('fill serial number done !!!! ')

    btn_repair = main_form.child_window(title="Repair", class_name="TBitBtn", found_index=0)
    btn_repair.click()
    print('click repair button done !!!! ')

    main_window = WaitForMainForm(timeout=10)
    repair_app  = Application(backend="win32").connect(handle=main_window.handle)
    repair_form = repair_app.window(handle=main_window.handle)
    repair_form.set_focus()
    time.sleep(1)

    red_codes = GetRedErrorCodesFromDBGrid(repair_form)

    if red_codes:
        print(f"\n✅ Red error codes: {red_codes}")
    else:
        print("\nℹ️ No red error codes found.")
        
if __name__ == "__main__":
    main()
