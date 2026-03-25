from pathlib import Path
import pandas as pd

def GetFormData():
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
            print("Data.xlsx has no data rows.")
            raise SystemExit

        missing_columns = [
            column_name for column_name in OUTPUT_COLUMN_MAP.values()
            if column_name not in df.columns
        ]

        if missing_columns:
            print("Missing columns in Data.xlsx:")
            for column in missing_columns:
                print(f"- {column}")
            raise SystemExit

        scanned_barcode = input("Scan Barcode: ").strip()
        # 25JF595448007P

        row = df.iloc[0]
        form_data = {
            "serial_number": scanned_barcode,
            "phenomenon": row["Pheomenon"],
            "failure_code": row["Failure Code"],
            "failure_desc": row["Failure Desc"],
            "location_code": row["Location Code"],
            "duty_code": row["Duty Code"],
            "reason_code": row["Reason Code"],
            "handling": row["Handling"],
            "duty_department": row["Duty Department"],
        }

        print(form_data)
        return form_data

    except PermissionError:
        print("Close Excel and try again.")
        return None


if __name__ == "__main__":
    print(GetFormData())
