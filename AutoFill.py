import time
import ctypes
import ReadData
import PIL.ImageGrab
from pywinauto import Application, Desktop

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

    if RedErrorExists(main_form):
        print("Success: Red Error found and clicked.")
    else:
        print("Notice: No Red Error found in the list.")

        
if __name__ == "__main__":
    main()
