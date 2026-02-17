<?php

namespace App\Services;

use App\Models\ExcelBatch;
use Illuminate\Support\Facades\Storage;


class ExcelBatchService
{
    public function create($file)
    {
        $path = $file->store('excel_batches');

        return ExcelBatch::create([
            'file_path' => $path,
            'original_name' => $file->getClientOriginalName()
        ]);
    }

    public function getPath(int $batchId): string
    {
        $batch = ExcelBatch::findOrFail($batchId);

        return Storage::path($batch->file_path);
    }
}
