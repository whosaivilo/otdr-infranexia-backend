<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    public function up(): void
    {
        Schema::table('otdr_histories', function (Blueprint $table) {
            $table->foreignId('file_upload_id')
                ->nullable()
                ->after('id')
                ->constrained('file_uploads')
                ->cascadeOnDelete();
        });
    }

    public function down(): void
    {
        Schema::table('otdr_histories', function (Blueprint $table) {
            $table->dropForeign(['file_upload_id']);
            $table->dropColumn('file_upload_id');
        });
    }
};
