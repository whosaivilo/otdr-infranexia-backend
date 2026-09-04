<?php

namespace App\Http\Controllers\Api;

use App\Http\Controllers\Controller;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Process;
use App\Models\OtdrHistory;

class OtdrController extends Controller
{
    public function upload(Request $request)
    {
        // 1. Validasi file masuk
        $request->validate([
            'file' => 'required|mimes:xlsx,xls|max:10240', // Max 10MB
        ]);

        // 2. Simpan file sementara di folder storage/app/public/otdr
        $path = $request->file('file')->store('public/otdr');
        $fullPath = storage_path('app/' . $path);

        // 3. Panggil script Python untuk mengolah file Excel
        $scriptPath = base_path('otdr_converter.py');

        // Catatan: Jika di Windows perintah 'python3' tidak dikenali, ganti menjadi 'python'
        $result = Process::run("python3 \"{$scriptPath}\" \"{$fullPath}\"");

        if ($result->failed()) {
            return response()->json([
                'success' => false,
                'message' => 'Gagal memproses file dengan Python.',
                'error' => $result->errorOutput()
            ], 500);
        }

        // 4. Tangkap hasil dari Python (berupa teks JSON) dan ubah jadi Array PHP
        $output = json_decode($result->output(), true);

        // 5. Simpan hasil kalkulasi ke database MySQL
        if (is_array($output)) {
            foreach ($output as $row) {
                // Pastikan key dari Python sama dengan nama kolom di database
                OtdrHistory::create($row);
            }
        }

        return response()->json([
            'success' => true,
            'message' => 'Kalkulasi OTDR Selesai',
            'data' => $output
        ], 200);
    }

    public function history()
    {
        // Ambil semua data riwayat dari database, urutkan dari yang terbaru
        $data = OtdrHistory::orderBy('created_at', 'desc')->get();

        return response()->json([
            'success' => true,
            'data' => $data
        ], 200);
    }
}
