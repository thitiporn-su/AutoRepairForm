from pathlib import Path
import pandas as pd

def GetFormData(scanned_barcode):
    OUTPUT_COLUMN_MAP = {
        "phenomenon": "Pheomenon",
        "failure_code": "Failure Code",
        "failure_desc": "Failure Desc",
        "location_code": "Location Code",
        "duty_code": "Duty Code",
        "reason_code": "Reason Code",
        "handling": "Handling",
        "duty_department": "Duty Department",
    }

    file_path = Path(__file__).with_name("Data.xlsx")

    try:
        df = pd.read_excel(file_path)
        df.columns = [str(column).strip() for column in df.columns]

        if df.empty:
            return None

        # ดึงแถวแรกมาใช้งาน (หรือจะเขียน Logic ค้นหาตามเงื่อนไขอื่นก็ได้)
        row = df.iloc[0]
        
        form_data = {
            "serial_number": scanned_barcode,
            "phenomenon": str(row["Pheomenon"]),
            "failure_code": str(row["Failure Code"]),
            "failure_desc": str(row["Failure Desc"]),
            "location_code": str(row["Location Code"]),
            "duty_code": str(row["Duty Code"]),
            "reason_code": str(row["Reason Code"]),
            "handling": str(row["Handling"]),
            "duty_department": str(row["Duty Department"]),
        }
        return form_data

    except Exception as e:
        print(f"Error reading Excel: {e}")
        return None


if __name__ == "__main__":
    print(GetFormData())
    # 25RO19392600R0