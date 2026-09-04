<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    public function up(): void
    {
        Schema::create('otdr_histories', function (Blueprint $table) {
            $table->id();
            $table->string('filename');
            $table->string('odc')->nullable();
            $table->integer('jumlah_titik_putus')->default(0);
            $table->integer('jumlah_bending')->default(0);
            $table->float('total_nilai_bending', 8, 4)->default(0);
            $table->float('estimasi_rx_onu', 8, 3)->default(0);
            $table->float('redaman_per_core', 8, 3)->default(0);
            $table->float('loss', 8, 3)->default(0);
            $table->float('threshold', 8, 3)->default(7.061);
            $table->string('fiber')->nullable();
            $table->string('wavelength')->nullable();
            $table->float('loss_db', 8, 3)->default(0);
            $table->float('length_km', 8, 4)->default(0);
            $table->float('attenuation', 8, 3)->default(0);
            $table->timestamps();
        });
    }

    public function down(): void
    {
        Schema::dropIfExists('otdr_histories');
    }
};
