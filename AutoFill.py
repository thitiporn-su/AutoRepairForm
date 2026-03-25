import time
import ctypes
import ReadData
import PIL.ImageGrab
from pywinauto import Application, Desktop

def RedErrorExists(window):
    rect = window.rectangle()
    img = PIL.ImageGrab.grab(bbox=(rect.left + 15, rect.top + 250, rect.right + 160, rect.bottom - 60))
    width, height = img.size
    print(f"Scanning Error List area {rect}")

    for y in range(0, height, 5):
        r, g, b = img.getpixel((30, y))

        if r > 200 and g < 50 and b < 50:
            print(f"Red pixel found at (30, {y}) with RGB({r}, {g}, {b})")
            window.click_input(coords=(30, y))
            print("Found Red Error.")
            return True
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

    window_handle = repair_windows[0].handle
    app = Application(backend="win32").connect(handle=window_handle)
    main_form = app.window(handle=window_handle)
    main_form.set_focus()

    # # When multiple matches exist, prefer the currently active top-level window.
    # foreground_handle = ctypes.windll.user32.GetForegroundWindow()
    # window = next((win for win in repair_windows if win.handle == foreground_handle), repair_windows[0])

    # app = Application(backend="win32").connect(handle=window.handle)
    # window.set_focus()

    # GetFormData = app.window(handle=window.handle)
    SN_input = main_form.child_window(class_name="TEdit", found_index=0)
    SN_input.set_edit_text(SN)
    print('fill serial number done !!!! ')

    btn_repair = main_form.child_window(title="Repair", class_name="TBitBtn", found_index=0)
    btn_repair.click()
    print('click repair button done !!!! ')

    # error_list = main_form.child_window(title_re=".*Error Code.*", visible_only=True)
    # parent_box = error_list.parent()
    # print(f"Found Label/Box: {error_list.window_text()} inside {parent_box.class_name()}")

    if RedErrorExists(main_form):
        print("Success: Red Error found and clicked.")
    else:
        print("Notice: No Red Error found in the list.")

        
if __name__ == "__main__":
    main()