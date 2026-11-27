<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\BelongsTo;
use Illuminate\Database\Eloquent\Relations\HasMany;

class FileUpload extends Model
{
    use HasFactory;

    protected $fillable = [
        'user_id',
        'original_filename',
        'raw_file_path',
        'output_file_path',
        'odc',
        'measurement_date',
    ];

    public function histories(): HasMany
    {
        return $this->hasMany(OtdrHistory::class, 'file_upload_id');
    }

    public function user(): BelongsTo
    {
        return $this->belongsTo(User::class, 'user_id');
    }
}
