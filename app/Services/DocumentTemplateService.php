<?php

namespace App\Services;

use Illuminate\Http\UploadedFile;
use Illuminate\Support\Facades\Log;
use Illuminate\Support\Facades\Storage;

Class DocumentTemplateService
{
    public function upTemplate(UploadedFile $file): string
    {
        $templateName = time() . '_' . $file->getClientOriginalName();

        Storage::putFileAs('templates', $file, $templateName);

        return $templateName;
    }

    public function getTemplates()
    {
        $files = Storage::files('templates');
        $templates = [];

        foreach ($files as $file) {
            $templates[] = basename($file);
        }

        return $templates;
    }

    public function getTemplate($templateName)
    {
        $path = "templates/{$templateName}";

        if (!Storage::exists($path)) {
            Log::error("Template no encontrado: " . $path);
            return null;
        }

        return Storage::path($path);
    }
}