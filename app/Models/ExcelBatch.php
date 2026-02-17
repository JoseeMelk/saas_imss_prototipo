<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use App\Models\ExcelData;

class ExcelBatch extends Model
{
    protected $fillable = [
        'file_path',
        'original_name',
        'is_processed',
    ];

    public function data()
    {
        return $this->hasMany(ExcelData::class, 'batch_id');
    }

    public function documents()
    {
        return $this->hasMany(Document::class, 'batch_id');
    }
}
