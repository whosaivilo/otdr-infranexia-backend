<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class OtdrHistory extends Model
{
    use HasFactory;

    // Properti ini sangat penting agar Laravel mengizinkan
    // data dari Python disimpan secara otomatis (Mass Assignment)
    protected $fillable = [
        'filename',
        'odc',
        'jumlah_titik_putus',
        'jumlah_bending',
        'total_nilai_bending',
        'estimasi_rx_onu',
        'redaman_per_core',
    ];
}
