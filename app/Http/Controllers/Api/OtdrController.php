<?php
namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use App\Models\OtdrHistory;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Process;

class OtdrController extends Controller
{
    public function upload(Request $request)
    {
        // 1. Validasi file masuk
        $request->validate([
            'file' => 'required|mimes:xlsx,xls|max:10240', // Max 10MB
        ]);

        // 2. Simpan file sementara di folder storage/app/public/otdr
        // Tambahkan 'public' agar Laravel 11 tahu kita menyimpannya di disk public
        $path = $request->file('file')->store('otdr', 'public');
        $fullPath = storage_path('app/public/' . $path);

        // 3. Panggil script Python untuk mengolah file Excel
        // File Python ada di root project (sejajar .env)
        $scriptPath = base_path('otdr_converter.py');

// Catatan: Jika di Windows perintah 'python3' tidak dikenali, ganti menjadi 'python'
        $result = Process::run("python3 \"{$scriptPath}\" \"{$fullPath}\"");

        $outputString = $result->output();
        $errorString  = $result->errorOutput();

        if ($result->failed()) {
            return response()->json([
                'success' => false,
                'message' => 'Gagal memproses file dengan Python.',
                // Tangkap error dari stderr maupun stdout
                'error'   => $errorString ?: $outputString,
            ], 500);
        }

        // 4. Tangkap hasil dari Python (berupa teks JSON) dan ubah jadi Array PHP
        $output = json_decode($outputString, true);

        // Cek error dari Python
        if (isset($output['error'])) {
            return response()->json([
                'success' => false,
                'message' => 'Python mengembalikan error.',
                'error'   => $output['error'],
            ], 500);
        }

        // 5. Simpan hasil kalkulasi ke database MySQL
        // Hanya simpan data per baris (rows), bukan summary
        if (isset($output['rows']) && is_array($output['rows'])) {
            foreach ($output['rows'] as $row) {
                OtdrHistory::create([
                    'filename'            => $row['filename'],
                    'odc'                 => $output['odc'] ?? ($row['odc'] ?? null),
                    'jumlah_titik_putus'  => $row['jumlah_titik_putus'] ?? 0,
                    'jumlah_bending'      => $row['jumlah_bending'] ?? 0,
                    'total_nilai_bending' => $row['total_nilai_bending'] ?? 0,
                    'estimasi_rx_onu'     => $row['estimasi_rx_onu'] ?? 0,
                    'redaman_per_core'    => $row['redaman_core'] ?? 0,
                    'loss'                => $row['loss'] ?? 0,
                    'threshold'           => $row['threshold'] ?? 7.0614781398215,
                    'fiber'               => $row['fiber'] ?? null,
                    'wavelength'          => $row['wavelength'] ?? null,
                    'loss_db'             => $row['loss_db'] ?? 0,
                    'length_km'           => $row['length_km'] ?? 0,
                    'attenuation'         => $row['attenuation'] ?? 0,
                ]);
            }
        }

        return response()->json([
            'success' => true,
            'message' => 'Kalkulasi OTDR Selesai',
            'data'    => $output,
        ], 200);
    }

    public function history()
    {
        // Ambil semua data riwayat dari database, urutkan dari yang terbaru
        $data = OtdrHistory::orderBy('created_at', 'desc')->get();

        return response()->json([
            'success' => true,
            'data'    => $data,
        ], 200);
    }
}
