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
        'loss',
        'threshold',
        'fiber',
        'wavelength',
        'loss_db',
        'length_km',
        'attenuation',
    ];

    protected $casts = [
        'jumlah_titik_putus' => 'integer',
        'jumlah_bending' => 'integer',
        'total_nilai_bending' => 'float',
        'estimasi_rx_onu' => 'float',
        'redaman_per_core' => 'float',
        'loss' => 'float',
        'threshold' => 'float',
        'loss_db' => 'float',
        'length_km' => 'float',
        'attenuation' => 'float',
    ];
}
