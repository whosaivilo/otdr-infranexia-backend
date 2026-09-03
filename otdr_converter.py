import sys
import json
import pandas as pd

def process_file(filepath):
    try:
        # Nanti teman TET kamu bisa menambahkan logika pandas untuk
        # membaca file excel mentah di sini: df = pd.read_excel(filepath)

        # Untuk saat ini, kita kembalikan dummy data yang sesuai dengan skema
        # database agar alur sistem (Upload -> Python -> Database -> React) bisa berjalan
        result = [
            {
                "filename": "FH-A-C1.sor",
                "odc": "ODC DUM FH",
                "jumlah_titik_putus": 0,
                "jumlah_bending": 1,
                "total_nilai_bending": 0.001342,
                "estimasi_rx_onu": -17.576,
                "redaman_per_core": 0.914
            },
            {
                "filename": "FH-A-C2.sor",
                "odc": "ODC DUM FH",
                "jumlah_titik_putus": 2,
                "jumlah_bending": 0,
                "total_nilai_bending": 1.298608,
                "estimasi_rx_onu": -17.733,
                "redaman_per_core": 1.258
            }
        ]

        # Cetak hasil sebagai JSON murni (akan ditangkap oleh fungsi Process di Laravel)
        print(json.dumps(result))

    except Exception as e:
        # Jika error, kirim balikan error agar mudah di-debug
        print(json.dumps({"error": str(e)}))
        sys.exit(1)

if __name__ == "__main__":
    # Menangkap argumen path file dari Laravel
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        process_file(file_path)
    else:
        print(json.dumps({"error": "Tidak ada file yang diberikan"}))
        sys.exit(1)

