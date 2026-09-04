import sys
import json
from openpyxl import load_workbook

# Konstanta default
RX_ONU_BASE = -16.0
DEFAULT_ODC = "ODC DUM FH"
DEFAULT_THRESHOLD = 7.0614781398215


def safe_float(val, default=0.0):
    """Konversi aman ke float, mengganti koma dengan titik."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(",", "."))
    except (ValueError, TypeError):
        return default


def parse_event_value(val):
    """
    Parse nilai event dari cell Excel.

    Returns:
        - "end"      : jika string "end" (case-insensitive)
        - "begin"    : jika string "begin" (case-insensitive)
        - float > 0  : jika angka (negatif diubah jadi positif / absolute)
        - None       : jika kosong atau tidak dikenali
    """
    if val is None:
        return None

    if isinstance(val, str):
        sval = val.strip().lower()
        if sval == "end":
            return "end"
        if sval == "begin":
            return "begin"
        # Coba parse angka
        try:
            return abs(float(sval.replace(",", ".")))
        except (ValueError, TypeError):
            return None

    if isinstance(val, (int, float)):
        return abs(float(val))

    return None


def read_input_file(filepath):
    """
    Membaca file Excel raw dan mengembalikan struktur data lengkap.
    """
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    # --- Metadata ---
    # Tanggal biasanya di baris 3, kolom 2
    date_val = ws.cell(row=3, column=2).value
    date_str = str(date_val).strip() if date_val else ""

    # --- Distance Headers ---
    # Di file raw, distance headers mulai dari baris 5, kolom 8
    distance_headers = []
    col_idx = 8
    while True:
        val = ws.cell(row=5, column=col_idx).value
        if val is None:
            break
        try:
            distance_headers.append(float(val))
        except (ValueError, TypeError):
            break
        col_idx += 1

    num_distance_cols = len(distance_headers)
    if num_distance_cols == 0:
        raise ValueError("Tidak ditemukan kolom jarak (mulai kolom 8 baris 5)")

    # --- Data Rows ---
    rows_data = []
    row_idx = 6
    while True:
        file_name = ws.cell(row=row_idx, column=2).value
        if file_name is None or str(file_name).strip() == "":
            break

        # Baca events per kolom jarak
        events = []
        for d_idx in range(num_distance_cols):
            col = 8 + d_idx
            cell_val = ws.cell(row=row_idx, column=col).value
            events.append(parse_event_value(cell_val))

        # Hitung redaman/core = jumlah absolute semua nilai numerik di baris ini
        redaman_core = sum(
            v for v in events if isinstance(v, (int, float))
        )

        # Baca nilai dasar
        attenuation = safe_float(ws.cell(row=row_idx, column=7).value)
        loss_db = safe_float(ws.cell(row=row_idx, column=5).value)
        length_km = safe_float(ws.cell(row=row_idx, column=6).value)

        # Estimasi RX ONU
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
            "loss": round(redaman_core, 3),   # Kolom Loss = Redaman/Core
            "events": events
        })

        row_idx += 1

    wb.close()

    return {
        "date": date_str,
        "distance_headers": distance_headers,
        "rows": rows_data
    }


def compute_summary(distance_headers, rows_data):
    """
    Hitung summary per kolom:
      - jumlah_titik_putus : hanya hitung "end"
      - jumlah_bending     : hitung semua nilai numerik
      - total_nilai_bending: distance_header + sum(abs(numeric_values))
    """
    num_cols = len(distance_headers)
    jumlah_titik_putus = []
    jumlah_bending = []
    total_nilai_bending = []

    for col_idx in range(num_cols):
        tp = 0
        jb = 0
        tnb = 0.0

        for row in rows_data:
            val = row["events"][col_idx]
            if val == "end":
                tp += 1
            elif isinstance(val, (int, float)):
                jb += 1
                tnb += val

        jumlah_titik_putus.append(tp)
        jumlah_bending.append(jb)
        total_nilai_bending.append(
            round(distance_headers[col_idx] + tnb, 13)
        )

    return {
        "jumlah_titik_putus": jumlah_titik_putus,
        "jumlah_bending_tipus": jumlah_bending,
        "total_nilai_bending": total_nilai_bending
    }


def process_file(filepath, odc_name=None, threshold=None):
    """
    Proses file Excel dan kembalikan dict lengkap siap di-JSON-kan.
    """
    raw_data = read_input_file(filepath)
    summary = compute_summary(raw_data["distance_headers"], raw_data["rows"])

    # Tambahkan threshold ke setiap row
    thr = threshold if threshold is not None else DEFAULT_THRESHOLD
    for row in raw_data["rows"]:
        row["threshold"] = thr

    return {
        "odc": odc_name if odc_name else DEFAULT_ODC,
        "date": raw_data["date"],
        "distance_headers": raw_data["distance_headers"],
        "summary": summary,
        "rows": raw_data["rows"]
    }


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({
            "error": "Penggunaan: python3 otdr_converter.py <file.xlsx> [odc_name] [threshold]"
        }))
        sys.exit(1)

    file_path = sys.argv[1]
    odc_name = sys.argv[2] if len(sys.argv) > 2 else None
    threshold = safe_float(sys.argv[3]) if len(sys.argv) > 3 else None

    try:
        result = process_file(file_path, odc_name, threshold)
        print(json.dumps(result, indent=2))
    except Exception as e:
        print(json.dumps({"error": f"Python Error: {str(e)}"}))
        sys.exit(1)
