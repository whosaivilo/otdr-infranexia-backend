import sys
import json
import re
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# === KONSTANTA ===
RX_ONU_BASE = -16.0
DEFAULT_ODC = "ODC DUM FH"
DEFAULT_THRESHOLD = 7.0614781398215

# === STYLING EXCEL (Berdasarkan Spesifikasi Baru) ===
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

# Kumpulan Font Arial sesuai permintaan
FONT_AR = Font(name="Arial", size=11)
FONT_AR_BOLD = Font(name="Arial", size=11, bold=True)
FONT_AR_WHITE_BOLD = Font(name="Arial", size=11, bold=True, color="FFFFFF")
FONT_AR_22_WHITE_BOLD = Font(name="Arial", size=22, bold=True, color="FFFFFF")
FONT_AR_36_BOLD = Font(name="Arial", size=36, bold=True)

# Kumpulan Warna Latar
FILL_BLACK = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
FILL_RED = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
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

    # 1. BACA HEADER ATAS SECARA DINAMIS (Baris 1 s/d 6)
    top_headers = []
    for r in range(1, 7):
        row_data = []
        for c in range(1, ws.max_column + 1):
            row_data.append(ws.cell(row=r, column=c).value)
        top_headers.append(row_data)

    date_val = top_headers[2][1] if len(top_headers[2]) > 1 else None
    date_str = str(date_val).strip() if date_val else datetime.now().strftime("%m/%d/%Y %H:%M:%S")

    # 2. BACA HEADER JARAK
    distance_headers = []
    col_idx = 8
    while True:
        val = ws.cell(row=7, column=col_idx).value if ws.cell(row=7, column=col_idx).value else ws.cell(row=5, column=col_idx).value
        if val is None or str(val).strip() == "": break
        try:
            distance_headers.append(float(str(val).replace(',', '.')))
        except: pass
        col_idx += 1

    # 3. BACA DATA INTI (Baris 8 ke bawah)
    rows_data = []
    row_idx = 8
    while True:
        file_name = ws.cell(row=row_idx, column=2).value
        if not file_name or str(file_name).strip() == "": break

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
    return {"date": date_str, "top_headers": top_headers, "distance_headers": distance_headers, "rows": rows_data}

def compute_summary(distance_headers, rows_data):
    num_cols = len(distance_headers)
    jumlah_tp, jumlah_bend, total_bend = [], [], []
    for col_idx in range(num_cols):
        tp, jb, tnb = 0, 0, 0.0
        for row in rows_data:
            val = row["events"][col_idx]
            if val == "end": tp += 1
            elif isinstance(val, (int, float)):
                jb += 1
                tnb += val
        jumlah_tp.append(tp)
        jumlah_bend.append(jb)
        total_bend.append(round(distance_headers[col_idx] + tnb, 13))
    return {"tp": jumlah_tp, "bend": jumlah_bend, "total": total_bend}

def create_formatted_excel(output_path, raw_data, summary, threshold):
    wb = Workbook()
    ws = wb.active
    ws.title = "Event Table"

    total_data_rows = len(raw_data["rows"])
    last_row_index = 7 + total_data_rows

    # === 1. TULIS HEADER ATAS (Baris 1 s/d 6) SECARA DINAMIS ===
    for r_idx, row_data in enumerate(raw_data["top_headers"], start=1):
        for c_idx, val in enumerate(row_data, start=1):
            if val is None: continue

            # Geser kolom ke kanan karena H dan I disisipkan
            target_col = c_idx if c_idx < 8 else c_idx + 2
            cell = ws.cell(row=r_idx, column=target_col, value=val)

            # Default Font
            cell.font = FONT_AR
            cell.alignment = CENTER_ALIGN

            str_val = str(val).strip().upper()

            # ATURAN FORMATTING DINAMIS BERDASARKAN KONTEN
            if r_idx == 4 and isinstance(val, (int, float)):
                # Angka besar di baris 4 (seperti 32, 27)
                cell.font = FONT_AR_36_BOLD

            elif re.match(r'^\d+(?:\.\d+)?M$', str_val):
                # Deteksi tulisan jarak otomatis (contoh: 150m, 200m)
                cell.font = FONT_AR_22_WHITE_BOLD
                cell.fill = FILL_RED

            elif "TITIK REPAIR" in str_val or "ODC" in str_val or "STO" in str_val:
                # Tulisan TITIK REPAIR / ODC / STO
                cell.font = FONT_AR_BOLD

            elif "TOTAL NILAI BENDING" in str_val:
                # Tulisan TOTAL NILAI BENDING di G4
                cell.font = FONT_AR_WHITE_BOLD
                cell.fill = FILL_BLACK

    # === 2. SUMMARY (Kolom I - Baris 1, 2, 3) ===
    headers_summary = ["JUMLAH TITIK PUTUS", "JUMLAH BENDING & TIPUS", "TOTAL NILAI BENDING"]
    for i, summary_text in enumerate(headers_summary, start=1):
        cell = ws.cell(row=i, column=9, value=summary_text)
        cell.font, cell.fill, cell.alignment = FONT_AR_WHITE_BOLD, FILL_BLACK, CENTER_ALIGN
        for d_idx in range(len(raw_data["distance_headers"])):
            v_cell = ws.cell(row=i, column=10 + d_idx, value=summary[["tp", "bend", "total"][i-1]][d_idx])
            v_cell.font, v_cell.alignment = FONT_AR, CENTER_ALIGN

    # === 3. RUMUS COUNTIF DI KOLOM H BARIS 4 ===
    h4_cell = ws.cell(row=4, column=8, value=f'=COUNTIF(H8:H{last_row_index}, "<-22")')
    h4_cell.font, h4_cell.alignment = FONT_AR_36_BOLD, CENTER_ALIGN

    # === 4. HEADER TABEL (Baris 7) ===
    headers_col = [(2,"File"), (3,"Fiber"), (4,"Wavelength"), (5,"Loss, dB"), (6,"Length, km"), (7,"Attenuation")]
    for col, val in headers_col:
        cell = ws.cell(row=7, column=col, value=val)
        cell.font, cell.alignment, cell.border = FONT_AR_BOLD, CENTER_ALIGN, THIN_BORDER

    # Header H dan I (ESTIMASI RX ONU & REDAMAN / CORE)
    for col, val in [(8, "ESTIMASI RX ONU"), (9, "REDAMAN / CORE")]:
        cell = ws.cell(row=7, column=col, value=val)
        cell.font, cell.fill, cell.alignment, cell.border = FONT_AR_WHITE_BOLD, FILL_BLACK, CENTER_ALIGN, THIN_BORDER

    # Header Jarak
    for d_idx, dist in enumerate(raw_data["distance_headers"]):
        cell = ws.cell(row=7, column=10 + d_idx, value=dist)
        cell.font, cell.alignment, cell.border = FONT_AR_BOLD, CENTER_ALIGN, THIN_BORDER

    # Header Kolom Terakhir (Loss & Threshold)
    loss_col = 10 + len(raw_data["distance_headers"])
    thresh_col = loss_col + 1
    for col, val in [(loss_col, "Loss"), (thresh_col, "Threshold")]:
        cell = ws.cell(row=7, column=col, value=val)
        cell.font, cell.fill, cell.alignment, cell.border = FONT_AR, FILL_YELLOW, CENTER_ALIGN, THIN_BORDER

    # === 5. ISI DATA & FORMATTING KUNING (Baris 8 ke bawah) ===
    current_row = 8
    for row in raw_data["rows"]:
        ws.cell(row=current_row, column=2, value=row["filename"]).font = FONT_AR
        ws.cell(row=current_row, column=3, value=row["fiber"]).font = FONT_AR
        ws.cell(row=current_row, column=4, value=row["wavelength"]).font = FONT_AR
        ws.cell(row=current_row, column=5, value=row["loss_db"]).font = FONT_AR
        ws.cell(row=current_row, column=6, value=row["length_km"]).font = FONT_AR
        ws.cell(row=current_row, column=7, value=row["attenuation"]).font = FONT_AR

        # Kolom H: Estimasi RX ONU (Kuning jika kurang dari -22)
        cell_rx = ws.cell(row=current_row, column=8, value=row["estimasi_rx_onu"])
        cell_rx.alignment, cell_rx.font = CENTER_ALIGN, FONT_AR
        if row["estimasi_rx_onu"] < -22.0:
            cell_rx.fill = FILL_YELLOW

        # Kolom I: Redaman Core
        ws.cell(row=current_row, column=9, value=row["redaman_core"]).font = FONT_AR

        # Event Jarak (Pencarian teks "end")
        for d_idx in range(len(raw_data["distance_headers"])):
            val = row["events"][d_idx]
            cell = ws.cell(row=current_row, column=10 + d_idx, value=val)
            cell.alignment = CENTER_ALIGN

            # Jika tulisan "end" -> Latar Kuning, Arial 11 Biasa (Tidak Bold)
            if str(val).lower() == "end":
                cell.font, cell.fill = FONT_AR, FILL_YELLOW
            else:
                cell.font = FONT_AR

        # Loss & Threshold data
        ws.cell(row=current_row, column=loss_col, value=row["loss"]).font = FONT_AR
        ws.cell(row=current_row, column=thresh_col, value=threshold).font = FONT_AR

        current_row += 1

    wb.save(output_path)
    wb.close()

# ================== PERUBAHAN DI SINI ==================
def process_file(filepath, output_path, odc_name=None, threshold=None):
    raw_data = read_input_file(filepath)
    summary = compute_summary(raw_data["distance_headers"], raw_data["rows"])
    thr = threshold if threshold is not None else DEFAULT_THRESHOLD
    for row in raw_data["rows"]: row["threshold"] = thr
    create_formatted_excel(output_path, raw_data, summary, thr)

    # KUNCI PERBAIKAN: Kembalikan data 'rows' ke Laravel agar React bisa menggambar tabel
    return {
        "odc": odc_name if odc_name else DEFAULT_ODC,
        "date": raw_data["date"],
        "rows": raw_data["rows"]
    }
# ======================================================

if __name__ == "__main__":
    if len(sys.argv) < 3: sys.exit(1)
    try:
        result = process_file(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
        print(json.dumps(result))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
