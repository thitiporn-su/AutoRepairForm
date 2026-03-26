import PIL.ImageGrab
import time
from pywinauto import Desktop, keyboard, mouse

# ─────────────────────────────────────────────────────────────────
#  FIELD ORDER  — top → bottom as they appear in the Repair Window
#  Set None to skip that row, or use the excel_data key to fill it
# ─────────────────────────────────────────────────────────────────
FIELD_ORDER = [
    "phenomenon",       # row 0  - Phenomenon
    "failure_code",     # row 1  - Failure Code
    None,               # row 2  - Failure Desc   (read-only)
    None,               # row 3  - Defect Area Code
    None,               # row 4  - LinkMonumber
    "location_code",    # row 5  - Location Code
    None,               # row 6  - Part No.
    None,               # row 7  - Part No.Desc
    None,               # row 8  - Vendor Code
    None,               # row 9  - Vendor Desc
    None,               # row 10 - Vendor D/C
    "duty_code",        # row 11 - Duty Code
    "reason_code",      # row 12 - Reason Code
    "handling",         # row 13 - Handling
    None,               # row 14 - Duty Department
]


# ─────────────────────────────────────────────────────────────────
#  PIXEL HELPERS
# ─────────────────────────────────────────────────────────────────

def IsInputColor(r, g, b):
    """
    Delphi input field backgrounds:
      White       r>240 g>240 b>240
      Light yellow r>230 g>225 b<210  (read-only / focused)
    """
    is_white  = r > 240 and g > 240 and b > 240
    is_yellow = r > 230 and g > 225 and b < 210
    return is_white or is_yellow


def FindInputFields(img, min_width=60, min_height=12, max_height=35):
    """
    Scan screenshot for rectangular input fields by background color.
    Returns list of (x_center, y_center, width, height) sorted top→bottom.
    """
    img_w, img_h = img.size
    pixels       = img.load()
    visited_rows = set()
    fields       = []

    for y in range(0, img_h - 2, 2):
        if y in visited_rows:
            continue

        run_start = None
        run_end   = None

        for x in range(0, img_w):
            r, g, b = pixels[x, y][:3]
            if IsInputColor(r, g, b):
                if run_start is None:
                    run_start = x
                run_end = x
            else:
                if run_start is not None and (x - run_start) > min_width:
                    break

        if run_start is None or run_end is None:
            continue
        run_width = run_end - run_start
        if run_width < min_width:
            continue

        # measure field height
        field_height = 0
        for dy in range(max_height + 1):
            check_y = y + dy
            if check_y >= img_h:
                break
            mid_x = run_start + run_width // 2
            r, g, b = pixels[mid_x, check_y][:3]
            if IsInputColor(r, g, b):
                field_height = dy + 1
            else:
                break

        if field_height < min_height:
            continue

        y_center = y + field_height // 2
        x_center = run_start + run_width // 2

        # skip duplicates within 10px vertically
        if any(abs(fy - y_center) < 10 for _, fy, _, _ in fields):
            continue

        fields.append((x_center, y_center, run_width, field_height))

        for dy in range(field_height + 2):
            visited_rows.add(y + dy)

    fields.sort(key=lambda f: f[1])
    return fields


# ─────────────────────────────────────────────────────────────────
#  MAIN FILL FUNCTION  (no OCR, no tesseract)
# ─────────────────────────────────────────────────────────────────

def FillRepairWindowByPixel(excel_data, log_fn=print):
    """
    Fills the Repair Window purely by scanning pixel colors.
    No pytesseract, no label search needed.
    """
    time.sleep(0.8)

    # 1. Find the Repair Window
    try:
        repair_win = Desktop(backend="win32").window(title="Repair Window")
        repair_win.wait("visible", timeout=10)
        repair_win.set_focus()
        time.sleep(0.3)
    except Exception as e:
        log_fn(f"Repair Window not found: {e}")
        return False

    win_rect = repair_win.rectangle()

    # 2. Screenshot the window
    img = PIL.ImageGrab.grab(bbox=(
        win_rect.left, win_rect.top,
        win_rect.right, win_rect.bottom
    ))
    img.save("debug_pixel_scan.png")  # open this to verify field detection

    # 3. Find all input fields by pixel color
    fields = FindInputFields(img)
    log_fn(f"Pixel scan found {len(fields)} input fields")
    for i, (fx, fy, fw, fh) in enumerate(fields):
        log_fn(f"  Field[{i}]  x={fx}  y={fy}  w={fw}  h={fh}")

    # 4. Fill each mapped field in order
    for idx, data_key in enumerate(FIELD_ORDER):
        if idx >= len(fields):
            log_fn(f"Ran out of detected fields at index {idx}")
            break

        if data_key is None:
            continue  # skip unmapped rows silently

        value = str(excel_data.get(data_key, "")).strip()
        if not value or value.lower() == "nan":
            log_fn(f"  Field[{idx}] '{data_key}' — empty, skipping")
            continue

        fx, fy, fw, fh = fields[idx]
        screen_x = win_rect.left + fx
        screen_y = win_rect.top  + fy

        mouse.click(button="left", coords=(screen_x, screen_y))
        time.sleep(0.15)
        keyboard.send_keys("^a",   pause=0.05)
        keyboard.send_keys(value,  with_spaces=True, pause=0.03)
        keyboard.send_keys("{TAB}", pause=0.2)

        log_fn(f"  Field[{idx}] '{data_key}' = {value}")

    log_fn("Pixel fill complete")
    return True