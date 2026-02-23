<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Services\DocumentTemplateService;

class DocumentTemplateController extends Controller
{
    protected $templateService;

    public function __construct(DocumentTemplateService $templateService)
    {
        $this->templateService = $templateService;
    }

    public function upload(Request $request)
    {
        $request->validate([
            'template' => 'required|file|mimes:doc,docx,odt',
        ]);

        $file = $request->file('template');
        $templateName = $this->templateService->upTemplate($file);

        return response()->json(['message' => 'Template uploaded successfully', 'template' => $templateName], 201);
    }

    public function index()
    {
        $templates = $this->templateService->getTemplates();
        return response()->json($templates);
    }

    public function show($templateName)
    {
        $templatePath = $this->templateService->getTemplate($templateName);

        if (!$templatePath) {
            return response()->json(['message' => 'Template not found'], 404);
        }

        return response()->download($templatePath, $templateName);
    }
}
