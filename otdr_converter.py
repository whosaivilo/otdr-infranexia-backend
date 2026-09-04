import sys
import json
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

# Konstanta default
RX_ONU_BASE = -16.0
DEFAULT_ODC = "ODC DUM FH"
DEFAULT_THRESHOLD = 7.0614781398215

def safe_float(val, default=0.0):
    if val is None: return default
    if isinstance(val, (int, float)): return float(val)
    try: return float(str(val).replace(",", "."))
    except (ValueError, TypeError): return default

def parse_event_value(val):
    if val is None: return None
    if isinstance(val, str):
        sval = val.strip().lower()
        if sval == "end": return "end"
        if sval == "begin": return "begin"
        try: return abs(float(sval.replace(",", ".")))
        except (ValueError, TypeError): return None
    if isinstance(val, (int, float)): return abs(float(val))
    return None

def read_input_file(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active
    date_val = ws.cell(row=3, column=2).value
    date_str = str(date_val).strip() if date_val else ""

    distance_headers = []
    col_idx = 8
    while True:
        val = ws.cell(row=5, column=col_idx).value
        if val is None: break
        try: distance_headers.append(float(val))
        except (ValueError, TypeError): break
        col_idx += 1

    num_distance_cols = len(distance_headers)
    if num_distance_cols == 0:
        raise ValueError("Tidak ditemukan kolom jarak")

    rows_data = []
    row_idx = 6
    while True:
        file_name = ws.cell(row=row_idx, column=2).value
        if file_name is None or str(file_name).strip() == "": break

        events = []
        for d_idx in range(num_distance_cols):
            val = parse_event_value(ws.cell(row=row_idx, column=8 + d_idx).value)
            events.append(val)

        redaman_core = sum(v for v in events if isinstance(v, (int, float)))
        attenuation = safe_float(ws.cell(row=row_idx, column=7).value)
        loss_db = safe_float(ws.cell(row=row_idx, column=5).value)
        length_km = safe_float(ws.cell(row=row_idx, column=6).value)
        estimasi_rx = RX_ONU_BASE - redaman_core - attenuation

        rows_data.append({
            "filename": str(file_name).strip(),
            "fiber": str(ws.cell(row=row_idx, column=3).value or "").strip(),
            "wavelength": str(ws.cell(row=row_idx, column=4).value or "").strip(),
            "loss_db": loss_db,
            "length_km": length_km,
            "attenuation": attenuation,
            "estimasi_rx_onu": round(estimasi_rx, 3),
            "redaman_core": round(redaman_core, 3),
            "loss": round(redaman_core, 3),
            "events": events
        })
        row_idx += 1
    wb.close()
    return {"date": date_str, "distance_headers": distance_headers, "rows": rows_data}

def compute_summary(distance_headers, rows_data):
    num_cols = len(distance_headers)
    jumlah_titik_putus, jumlah_bending, total_nilai_bending = [], [], []

    for col_idx in range(num_cols):
        tp, jb, tnb = 0, 0, 0.0
        for row in rows_data:
            val = row["events"][col_idx]
            if val == "end": tp += 1
            elif isinstance(val, (int, float)):
                jb += 1
                tnb += val
        jumlah_titik_putus.append(tp)
        jumlah_bending.append(jb)
        total_nilai_bending.append(round(distance_headers[col_idx] + tnb, 13))

    return {
        "jumlah_titik_putus": jumlah_titik_putus,
        "jumlah_bending_tipus": jumlah_bending,
        "total_nilai_bending": total_nilai_bending
    }

# FUNGSI BARU: Membuat File Excel yang Rapi
def create_formatted_excel(output_path, raw_data, summary, threshold):
    wb = Workbook()
    ws = wb.active
    ws.title = "Event Table"

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    header_font = Font(bold=True, size=10, name="Calibri")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    data_font = Font(size=10, name="Calibri")

    # Baris Summary (1-3)
    ws.cell(row=1, column=9, value="JUMLAH TITIK PUTUS").font = header_font
    ws.cell(row=2, column=9, value="JUMLAH BENDING & TIPUS").font = header_font
    ws.cell(row=3, column=9, value="TOTAL NILAI BENDING").font = header_font
    for i in range(1, 4): ws.cell(row=i, column=9).alignment = center_align

    for d_idx in range(len(raw_data["distance_headers"])):
        col = 10 + d_idx
        ws.cell(row=1, column=col, value=summary["jumlah_titik_putus"][d_idx]).alignment = center_align
        ws.cell(row=2, column=col, value=summary["jumlah_bending_tipus"][d_idx]).alignment = center_align
        ws.cell(row=3, column=col, value=summary["total_nilai_bending"][d_idx]).alignment = center_align

    # Baris Header Kolom (Baris 7)
    headers = ["File", "Fiber", "Wavelength", "Loss, dB", "Length, km", "Attenuation, dB/km", "ESTIMASI RX ONU", "REDAMAN / CORE"]
    for col_idx, val in enumerate(headers, start=2):
        cell = ws.cell(row=7, column=col_idx, value=val)
        cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, center_align, thin_border

    for d_idx, dist in enumerate(raw_data["distance_headers"]):
        cell = ws.cell(row=7, column=10 + d_idx, value=dist)
        cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, center_align, thin_border

    loss_col = 10 + len(raw_data["distance_headers"])
    thresh_col = loss_col + 1
    for col_idx, val in [(loss_col, "Loss"), (thresh_col, "Threshold")]:
        cell = ws.cell(row=7, column=col_idx, value=val)
        cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, center_align, thin_border

    # Data Baris (Mulai dari Baris 8)
    current_row = 8
    for row in raw_data["rows"]:
        ws.cell(row=current_row, column=2, value=row["filename"]).font = data_font
        ws.cell(row=current_row, column=3, value=row["fiber"]).font = data_font
        ws.cell(row=current_row, column=4, value=row["wavelength"]).font = data_font
        ws.cell(row=current_row, column=5, value=row["loss_db"]).font = data_font
        ws.cell(row=current_row, column=6, value=row["length_km"]).font = data_font
        ws.cell(row=current_row, column=7, value=row["attenuation"]).font = data_font
        ws.cell(row=current_row, column=8, value=row["estimasi_rx_onu"]).font = data_font
        ws.cell(row=current_row, column=9, value=row["redaman_core"]).font = data_font

        for d_idx in range(len(raw_data["distance_headers"])):
            val = row["events"][d_idx]
            cell = ws.cell(row=current_row, column=10 + d_idx, value=val)
            if val == "end": cell.font = Font(size=10, name="Calibri", color="FF0000")
            else: cell.font = data_font
            cell.alignment = center_align

        ws.cell(row=current_row, column=loss_col, value=row["loss"]).font = data_font
        ws.cell(row=current_row, column=thresh_col, value=threshold).font = data_font
        current_row += 1

    wb.save(output_path)
    wb.close()

def process_file(filepath, output_path, odc_name=None, threshold=None):
    raw_data = read_input_file(filepath)
    summary = compute_summary(raw_data["distance_headers"], raw_data["rows"])
    thr = threshold if threshold is not None else DEFAULT_THRESHOLD
    for row in raw_data["rows"]: row["threshold"] = thr

    # Buat File Excel
    create_formatted_excel(output_path, raw_data, summary, thr)

    return {
        "odc": odc_name if odc_name else DEFAULT_ODC,
        "date": raw_data["date"],
        "distance_headers": raw_data["distance_headers"],
        "summary": summary,
        "rows": raw_data["rows"]
    }

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Penggunaan: python3 otdr_converter.py <input.xlsx> <output.xlsx>"}))
        sys.exit(1)

    file_path = sys.argv[1]
    out_path = sys.argv[2]
    odc_name = sys.argv[3] if len(sys.argv) > 3 else None
    threshold = safe_float(sys.argv[4]) if len(sys.argv) > 4 else None

    try:
        result = process_file(file_path, out_path, odc_name, threshold)
        print(json.dumps(result))
    except Exception as e:
        print(json.dumps({"error": f"Python Error: {str(e)}"}))
        sys.exit(1)
