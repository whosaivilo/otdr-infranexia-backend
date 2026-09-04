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
            $table->timestamps();
        });
    }

    public function down(): void
    {
        Schema::dropIfExists('otdr_histories');
    }
};
