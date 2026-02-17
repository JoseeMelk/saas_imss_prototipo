<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class Document extends Model
{
    protected $fillable = [
        'name',
        'path',
        'pc',
        'ic',
        'ac',
        'has_images',
        'is_completed',
        'row_id',
        'batch_id',
    ];

    public function batch()
    {
        return $this->belongsTo(ExcelBatch::class, 'batch_id');
    }
}
