<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use App\Models\ExcelBatch;

class ExcelData extends Model
{
    protected $fillable = [
        'sheet',
        'columns',
        'datas',
        'batch_id'
    ];

    public function batch()
    {
        return $this->belongsTo(ExcelBatch::class, 'batch_id');
    }
}
