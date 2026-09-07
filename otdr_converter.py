import sys
import json
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# === KONSTANTA ===
RX_ONU_BASE = -16.0
DEFAULT_ODC = "ODC DUM FH"
DEFAULT_THRESHOLD = 7.0614781398215

# === STYLING EXCEL ===
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

FONT_AR = Font(name="Arial", size=11)
FONT_AR_BOLD = Font(name="Arial", size=11, bold=True)
FONT_AR_WHITE_BOLD = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FONT_AR_22_WHITE_BOLD = Font(name="Arial", size=22, bold=True, color="FFFFFF")
FONT_AR_22_BLACK_BOLD = Font(name="Arial", size=22, bold=True, color="000000")
FONT_AR_20_BOLD = Font(name="Arial", size=20, bold=True)
FONT_AR_36_BOLD = Font(name="Arial", size=36, bold=True)

FILL_BLACK = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
FILL_RED = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
FILL_HEADER_GRAY = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

def safe_float(val, default=0.0):
    if val is None: return default
    if isinstance(val, (int, float)): return float(val)
    try: return float(str(val).replace(",", "."))
    except: return default

def parse_event_value(val):
    if val is None: return None
    if isinstance(val, str):
        sval = val.strip().lower()
        if sval == "end": return "end"
        if sval == "begin": return "begin"
        try: return abs(float(sval.replace(",", ".")))
        except: return None
    if isinstance(val, (int, float)): return abs(float(val))
    return None

def read_input_file(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    dist_row = 7
    for r in range(1, 15):
        val = ws.cell(row=r, column=2).value
        if val and str(val).strip().lower() == "file":
            dist_row = r
            break

    date_val = ws.cell(row=max(1, dist_row - 2), column=2).value
    date_str = str(date_val).strip() if date_val else datetime.now().strftime("%m/%d/%Y %H:%M:%S")

    distance_headers = []
    col_idx = 8
    while True:
        val = ws.cell(row=dist_row, column=col_idx).value
        if val is None or str(val).strip() == "":
            val = ws.cell(row=5, column=col_idx).value

        if val is None or str(val).strip() == "":
            break

        try:
            distance_headers.append(float(str(val).replace(',', '.')))
        except:
            pass
        col_idx += 1

    rows_data = []
    row_idx = dist_row + 1
    while True:
        file_name = ws.cell(row=row_idx, column=2).value
        if not file_name or str(file_name).strip() == "":
            break

        events = []
        for d_idx in range(len(distance_headers)):
            events.append(parse_event_value(ws.cell(row=row_idx, column=8 + d_idx).value))

        redaman_core = sum(v for v in events if isinstance(v, (int, float)))
        attenuation = safe_float(ws.cell(row=row_idx, column=7).value)
        estimasi_rx = RX_ONU_BASE - redaman_core - attenuation

        rows_data.append({
            "filename": str(file_name).strip(),
            "fiber": str(ws.cell(row=row_idx, column=3).value or "").strip(),
            "wavelength": str(ws.cell(row=row_idx, column=4).value or "").strip(),
            "loss_db": safe_float(ws.cell(row=row_idx, column=5).value),
            "length_km": safe_float(ws.cell(row=row_idx, column=6).value),
            "attenuation": attenuation,
            "estimasi_rx_onu": round(estimasi_rx, 3),
            "redaman_core": round(redaman_core, 3),
            "loss": round(redaman_core, 3),
            "events": events
        })
        row_idx += 1
    wb.close()
    return {"date": date_str, "distance_headers": distance_headers, "rows": rows_data}

def create_formatted_excel(output_path, raw_data, threshold):
    wb = Workbook()
    ws = wb.active
    ws.title = "Event Table"

    total_data_rows = len(raw_data["rows"])
    last_row_index = 7 + total_data_rows if total_data_rows > 0 else 8

    # === 1. TULIS HEADER ATAS DENGAN INJEKSI RUMUS EXCEL ===
    c_tp = ws.cell(row=1, column=9, value="JUMLAH TITIK PUTUS")
    c_tp.font, c_tp.fill, c_tp.alignment = FONT_AR_WHITE_BOLD, FILL_BLACK, CENTER_ALIGN
    for d_idx in range(len(raw_data["distance_headers"])):
        col_letter = get_column_letter(10 + d_idx)
        vc = ws.cell(row=1, column=10 + d_idx, value=f'=COUNTIF({col_letter}8:{col_letter}{last_row_index}, "end")')
        vc.font, vc.alignment = FONT_AR, CENTER_ALIGN

    c_jb = ws.cell(row=2, column=9, value="JUMLAH BENDING & TIPUS")
    c_jb.font, c_jb.fill, c_jb.alignment = FONT_AR_WHITE_BOLD, FILL_BLACK, CENTER_ALIGN
    for d_idx in range(len(raw_data["distance_headers"])):
        col_letter = get_column_letter(10 + d_idx)
        vc = ws.cell(row=2, column=10 + d_idx, value=f'=COUNT({col_letter}8:{col_letter}{last_row_index})')
        vc.font, vc.alignment = FONT_AR, CENTER_ALIGN
        if (10 + d_idx) in [17, 24]:
            vc.font = FONT_AR_20_BOLD

    ws.cell(row=3, column=4, value="kabel 264").font = FONT_AR
    c_tn = ws.cell(row=3, column=9, value="TOTAL NILAI BENDING")
    c_tn.font, c_tn.fill, c_tn.alignment = FONT_AR_WHITE_BOLD, FILL_BLACK, CENTER_ALIGN
    for d_idx in range(len(raw_data["distance_headers"])):
        col_letter = get_column_letter(10 + d_idx)
        vc = ws.cell(row=3, column=10 + d_idx, value=f'=SUMIF({col_letter}7:{col_letter}{last_row_index}, "<>end")')
        vc.font, vc.alignment = FONT_AR, CENTER_ALIGN

    # Baris 4
    ws.cell(row=4, column=1, value="STO").font = FONT_AR_BOLD
    ws.cell(row=4, column=2, value="ODC DUM FH").font = FONT_AR_BOLD
    ws.cell(row=4, column=4, value="8 km").font = FONT_AR
    ws.cell(row=4, column=5, value="Panjang kabel 9,.").font = FONT_AR

    c_t = ws.cell(row=4, column=7, value="TOTAL NILAI BENDING")
    c_t.font, c_t.fill, c_t.alignment = FONT_AR_WHITE_BOLD, FILL_BLACK, CENTER_ALIGN

    c_h4 = ws.cell(row=4, column=8, value=f'=COUNTIF(H8:H{last_row_index}, "<-22")')
    c_h4.font, c_h4.alignment = FONT_AR_36_BOLD, CENTER_ALIGN

    c_200 = ws.cell(row=4, column=12, value="200m")
    c_200.font, c_200.alignment = FONT_AR_22_BLACK_BOLD, CENTER_ALIGN
    ws.cell(row=4, column=13).fill = FILL_YELLOW

    c_q4 = ws.cell(row=4, column=17, value=2)
    c_q4.font, c_q4.alignment = FONT_AR_36_BOLD, CENTER_ALIGN

    c_x4 = ws.cell(row=4, column=24, value=2)
    c_x4.font, c_x4.alignment = FONT_AR_36_BOLD, CENTER_ALIGN

    c_250 = ws.cell(row=4, column=32, value="250m")
    c_250.font, c_250.alignment = FONT_AR_22_BLACK_BOLD, CENTER_ALIGN
    ws.cell(row=4, column=33).fill = FILL_YELLOW

    ws.cell(row=4, column=34, value="kabel 48").font = FONT_AR

    # Baris 5
    ws.cell(row=5, column=12).fill = FILL_RED
    ws.cell(row=5, column=32).fill = FILL_RED
    ws.cell(row=5, column=14, value="kabel 264").font = FONT_AR

    c_150_1 = ws.cell(row=5, column=17, value="150m")
    c_150_1.font, c_150_1.fill, c_150_1.alignment = FONT_AR_22_WHITE_BOLD, FILL_RED, CENTER_ALIGN
    ws.cell(row=5, column=18, value="kabel 264").font = FONT_AR

    c_150_2 = ws.cell(row=5, column=24, value="150m")
    c_150_2.font, c_150_2.fill, c_150_2.alignment = FONT_AR_22_WHITE_BOLD, FILL_RED, CENTER_ALIGN

    ws.cell(row=5, column=25, value="kabel 264").font = FONT_AR
    ws.cell(row=5, column=34, value="TITIK").font = FONT_AR_BOLD

    # Baris 6
    ws.cell(row=6, column=1, value="Date:").font = FONT_AR
    ws.cell(row=6, column=2, value=raw_data["date"]).font = FONT_AR
    for col in [12, 17, 24, 32]:
        ws.cell(row=6, column=col, value="TITIK REPAIR").font = FONT_AR_BOLD
    ws.cell(row=6, column=34, value="ODC").font = FONT_AR_BOLD

    # === 2. HEADER TABEL UTAMA ===
    for col, val in [(2,"File"), (3,"Fiber"), (4,"Wavelength"), (5,"Loss, dB"), (6,"Length, km"), (7,"Attenuation")]:
        c = ws.cell(row=7, column=col, value=val)
        c.font, c.alignment, c.border = FONT_AR_BOLD, CENTER_ALIGN, THIN_BORDER

    for col, val in [(8, "ESTIMASI RX ONU"), (9, "REDAMAN / CORE")]:
        c = ws.cell(row=7, column=col, value=val)
        c.font, c.fill, c.alignment, c.border = FONT_AR_WHITE_BOLD, FILL_BLACK, CENTER_ALIGN, THIN_BORDER

    for d_idx, dist in enumerate(raw_data["distance_headers"]):
        c = ws.cell(row=7, column=10 + d_idx, value=dist)
        c.font, c.alignment, c.border = FONT_AR_BOLD, CENTER_ALIGN, THIN_BORDER

    loss_col = 10 + len(raw_data["distance_headers"])
    thresh_col = loss_col + 1
    for col, val in [(loss_col, "Loss"), (thresh_col, "Threshold")]:
        c = ws.cell(row=7, column=col, value=val)
        c.font, c.fill, c.alignment, c.border = FONT_AR_BOLD, FILL_YELLOW, CENTER_ALIGN, THIN_BORDER

    # === 3. ISI DATA (DENGAN LOGIKA CONDITIONAL FORMATTING BARU) ===
    current_row = 8
    for row in raw_data["rows"]:
        # 3a. Kolom Basic (Tanpa kolom 8 dan 9 karena butuh custom logic)
        for col, key in [(2,"filename"), (3,"fiber"), (4,"wavelength"), (5,"loss_db"), (6,"length_km"), (7,"attenuation"), (loss_col,"loss"), (thresh_col,"threshold")]:
            ws.cell(row=current_row, column=col, value=row[key]).font = FONT_AR

        # 3b. Kolom 8: Estimasi RX ONU (Kuning jika < -22.0)
        cr_rx = ws.cell(row=current_row, column=8, value=row["estimasi_rx_onu"])
        cr_rx.font, cr_rx.alignment = FONT_AR, CENTER_ALIGN
        if row["estimasi_rx_onu"] < -22.0:
            cr_rx.fill = FILL_YELLOW

        # 3c. Kolom 9: Redaman / Core (Kuning jika > 6)
        cr_redaman = ws.cell(row=current_row, column=9, value=row["redaman_core"])
        cr_redaman.font = FONT_AR
        if row["redaman_core"] > 6.0:
            cr_redaman.fill = FILL_YELLOW

        # 3d. Kolom 10 s/d Habis: Data Events (Kuning jika "end" ATAU >= 0.75)
        for d_idx in range(len(raw_data["distance_headers"])):
            val = row["events"][d_idx]
            ce = ws.cell(row=current_row, column=10 + d_idx, value=val)
            ce.alignment = CENTER_ALIGN
            ce.font = FONT_AR

            # Pengecekan Syarat Kuning
            if str(val).lower() == "end":
                ce.fill = FILL_YELLOW
            elif isinstance(val, (int, float)) and val >= 0.75:
                ce.fill = FILL_YELLOW

        current_row += 1

    wb.save(output_path)
    wb.close()

def process_file(filepath, output_path, odc_name=None, threshold=None):
    raw_data = read_input_file(filepath)
    thr = threshold if threshold is not None else DEFAULT_THRESHOLD
    for row in raw_data["rows"]: row["threshold"] = thr
    create_formatted_excel(output_path, raw_data, thr)

    return {
        "odc": odc_name if odc_name else DEFAULT_ODC,
        "date": raw_data["date"],
        "rows": raw_data["rows"]
    }

if __name__ == "__main__":
    if len(sys.argv) < 3: sys.exit(1)
    try:
        result = process_file(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
        print(json.dumps(result))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
