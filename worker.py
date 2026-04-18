import pandas as pd
import math
from datetime import time

def parse_hhmm(hhmm: str) -> float:
    try:
        hh, mm = hhmm.split(":")
        hh = int(hh)
        mm = int(mm)
        return hh + (mm / 60)
    except:
        return 0.0


def autosize_columns_xlsxwriter(worksheet, df):
    for idx, col in enumerate(df.columns):
        # หา max length ของทั้ง column
        max_len = max(
            df[col].astype(str).map(len).max(),
            len(col)
        ) + 2
        worksheet.set_column(idx, idx, max_len)

def process(input_path, output_path,
            include_data=False, no_pairing=False, keep_unpaired=False,
            base_day="16:00", offset_day=3,
            base_night="00:00", offset_night=3):
    
    df = pd.read_excel(input_path)

    base_day = parse_hhmm(base_day)         # → float
    base_night = parse_hhmm(base_night)     # → float
    
    # ---------------------------
    # 1) Mapping columns by index
    # ---------------------------
    df['supplier'] = df.iloc[:, 7]
    df['emp_code'] = df.iloc[:, 8]
    df['emp_name'] = df.iloc[:, 9]
    df['datetime'] = df.iloc[:, 15]
    df['device'] = df.iloc[:, 17]

    df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')

    # ---------------------------
    # 2) Mark IN / OUT
    # ---------------------------
    df['scan'] = df['device'].map({
        'ADMIN99900401': 'IN',
        'ADMIN99900402': 'OUT'
    })

    df_in = df[df['scan'] == 'IN'].sort_values(['emp_code', 'datetime'])
    df_out = df[df['scan'] == 'OUT'].sort_values(['emp_code', 'datetime'])

    paired = []
    # Group ตามพนักงานเพื่อจับคู่ทีละคน
    df_grouped = df.groupby("emp_code")

    for emp, rows in df_grouped:
        rows = rows.sort_values("datetime").reset_index(drop=True)
        used = set()
        max_hours = 24
        
        for i, row in rows.iterrows():
            if i in used:
                continue

            if row["scan"] == "IN":

                out_rows = rows[
                    (rows["scan"] == "OUT") &
                    (rows.index > i) &
                    (~rows.index.isin(used)) &
                    (rows["datetime"] > row["datetime"]) &
                    ((rows["datetime"] - row["datetime"]).dt.total_seconds() <= max_hours * 3600)
                ]

                if len(out_rows) > 0:
                    row_out = out_rows.iloc[0]
                    used.add(row_out.name)   # Mark ว่า OUT ตัวนี้ถูกใช้แล้ว

                    paired.append({
                        "supplier": row.get("supplier",""),
                        "emp_code": emp,
                        "emp_name": row.get("emp_name",""),
                        "time_in": row["datetime"],
                        "time_out": row_out["datetime"],
                        "status": "ปกติ"
                    })
                else:
                    if keep_unpaired:
                        paired.append({
                            "supplier": row.get("supplier",""),
                            "emp_code": emp,
                            "emp_name": row.get("emp_name",""),
                            "time_in": row["datetime"],
                            "time_out": pd.NaT,
                            "status": "ไม่พบเวลาออก"
                        })

            elif row["scan"] == "OUT":
                if i in used:
                    continue

                in_rows = rows[
                    (rows["scan"] == "IN") &
                    (rows.index < i) &
                    (~rows.index.isin(used)) &    # กันไม่ให้ผูก IN ที่ถูกใช้งานแล้ว
                    (rows["datetime"] < row["datetime"])
                ]

                if len(in_rows) == 0 and keep_unpaired:
                    paired.append({
                        "supplier": row.get("supplier",""),
                        "emp_code": emp,
                        "emp_name": row.get("emp_name",""),
                        "time_in": pd.NaT,
                        "time_out": row["datetime"],
                        "status": "ไม่พบเวลาเข้า"
                    })
                    used.add(i)



    result = pd.DataFrame(paired)

    # ---------------------------
    # Fill supplier blank as Part-Time
    # ---------------------------
    result['supplier'] = result['supplier'].astype(str).str.strip()
    result['supplier'] = result['supplier'].replace(["", "nan", "None"], "พนักงานบริษัท NuiB (Part-Time)")
    result['supplier'] = result['supplier'].fillna("พนักงานบริษัท NuiB (Part-Time)")

    # ---------------------------
    # Add scan info
    # ---------------------------
    result['scan_in'] = 'ADMIN99900401'
    result['scan_out'] = 'ADMIN99900402'

    # ---------------------------
    # Work hours & OT
    # ---------------------------
    result['work_hours'] = (result['time_out'] - result['time_in']).dt.total_seconds() / 3600

    # ถ้าเป็น NaN ให้คงค่า NaN
    result['work_hours'] = result['work_hours'].apply(lambda x: math.floor(x) if pd.notna(x) else x)

    def classify_shift(dt):
        t = dt.time()

        # --- กะบ่าย ---
        bh = int(base_day)
        bm = int((base_day - bh) * 60)

        day_center = dt.replace(hour=bh, minute=bm)
        start_day = (day_center - pd.Timedelta(hours=offset_day)).time()
        end_day   = (day_center + pd.Timedelta(hours=offset_day)).time()

        if start_day <= t <= end_day:
            return "บ่าย (Pick up-ขาเข้า)"

        # --- กะดึก ---
        nh = int(base_night)
        nm = int((base_night - nh) * 60)

        night_center = dt.replace(hour=nh, minute=nm)
        start_night = (night_center - pd.Timedelta(hours=offset_night)).time()
        end_night   = (night_center + pd.Timedelta(hours=offset_night)).time()

        # กรณีข้ามวัน เช่น 00:00 - 04:00
        if start_night <= t or t <= end_night:
            return "ดึก (Delivery-ขาออก)"

        return "ผิดปกติ"


    def safe_classify(dt):
        if pd.isna(dt):
            return 'ไม่พบเวลาเข้า'
        return classify_shift(dt)

    result['shift'] = result['time_in'].apply(safe_classify)

    result['shift'] = result['shift'].apply(
        lambda x: x if x in [
            'บ่าย (Pick up-ขาเข้า)',
            'ดึก (Delivery-ขาออก)'
        ] else 'ผิดปกติ'
    )

    result['ot_hours'] = result['work_hours'] - 9
    result['ot_hours'] = result['ot_hours'].apply(
        lambda x: max(min(x, 5), 0) if pd.notna(x) else x
    )

    result['work_hours'] = result['work_hours'].apply(lambda x: x if 9 <= x <= 16 else 'ผิดปกติ')

    def status_row(row):
        if pd.isna(row['work_hours']):
            return row['status'] if 'status' in row else 'ไม่สมบูรณ์'
        if row['shift'] not in ['บ่าย (Pick up-ขาเข้า)','ดึก (Delivery-ขาออก)']:
            return 'ผิดปกติ'
        if row['work_hours'] == 'ผิดปกติ':
            return 'ผิดปกติ'
        return 'ปกติ'


    result['สถานะ'] = result.apply(status_row, axis=1)

    # ---------------------------
    # Split date/time
    # ---------------------------
    result['วันที่ทำงาน'] = result['time_in'].dt.date.astype(str)
    result['เวลาทำงาน'] = result['time_in'].dt.time
    result['วันที่ออกงาน'] = result['time_out'].dt.date.astype(str)
    result['เวลาออกงาน'] = result['time_out'].dt.time

    result.drop(columns=['time_in', 'time_out'], inplace=True)

    # ---------------------------
    # Rename columns
    # ---------------------------
    result = result.rename(columns={
        'supplier': 'ชื่อซัพพลายเออร์',
        'emp_code': 'รหัสพนักงาน',
        'emp_name': 'ชื่อพนักงาน',
        'scan_in': 'สแกนเข้า',
        'scan_out': 'สแกนออก',
        'shift': 'กะงาน',
        'work_hours': 'ชั่วโมงทำงาน',
        'ot_hours': 'ชั่วโมง OT'
    })

    # ---------------------------
    # Order columns
    # ---------------------------
    result = result[[
        'ชื่อซัพพลายเออร์',
        'รหัสพนักงาน',
        'ชื่อพนักงาน',
        'สแกนเข้า',
        'สแกนออก',
        'วันที่ทำงาน',
        'เวลาทำงาน',
        'วันที่ออกงาน',
        'เวลาออกงาน',
        'กะงาน',
        'ชั่วโมงทำงาน',
        'ชั่วโมง OT',
        'สถานะ'
    ]]

    # ---------------------------
    # Export excel
    # ---------------------------
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        result.to_excel(writer, sheet_name='Result', index=False)

         # ถ้าเลือกให้ export sheet DATA
        if include_data:
            raw = pd.read_excel(input_path)
            raw.to_excel(writer, sheet_name='DATA', index=False)
            
            worksheet = writer.sheets['DATA']
            autosize_columns_xlsxwriter(worksheet, raw)

        workbook = writer.book
        worksheet = writer.sheets['Result']
        left_fmt = workbook.add_format({'align':'left','valign':'vcenter','text_wrap':False})
        worksheet.autofilter(0, 0, len(result), len(result.columns)-1)

        for i, col in enumerate(result.columns):
            max_len = max(result[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, max_len, left_fmt)

        for row in range(1, len(result) + 1):
            worksheet.set_row(row, 15)