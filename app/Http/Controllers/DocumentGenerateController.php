<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Storage;
use App\Services\DocumentTemplateService;
use App\Services\DocumentGenerateService;

class DocumentGenerateController extends Controller
{
    protected $templateService;
    protected $generateService;

    public function __construct(DocumentTemplateService $templateService, DocumentGenerateService $generateService)
    {
        $this->templateService = $templateService;
        $this->generateService = $generateService;
    }

    /**
     * Ejemplo de data
     * {
     *   "template": "mi_plantilla.docx",
     *   "data": {
     *     "ic": "Los alimentos no cumplen con las condiciones de almacenamiento ya que se encontro crema acida  en recipiente con capacidad muy  grande para el consumo diario el cual esta abierto para su utilización desde el 01/09/2025.	",
     *     "ac": "Se verificó que la directora implementara estrategias y capacitara al personal correspondiente para garantizar el correcto almacenamiento de los alimentos.",
     *  }
     */
    public function generate(Request $request)
    {
        $request->validate([
            'template' => 'required|string',
            'data' => 'required|array',
        ]);

        $templateName = $request->input('template');
        $data = $request->input('data');

        $templatePath = $this->templateService->getTemplate($templateName);

        if (!$templatePath) {
            return response()->json(['message' => 'Template not found'], 404);
        }

        // 🔹 nombre y ruta relativa
        $filename = time() . '_generated.docx';
        $relativePath = "generated/$filename";

        // 🔹 generar documento usando storage
        $this->generateService->generateDocument(
            $templatePath,
            $data,
            $relativePath
        );

        // 🔹 descargar desde storage
        return Storage::download($relativePath);
    }

    /**
     * Agrega imágenes a un documento generado previamente.
     *
     * Body (multipart/form-data):
     * {
     *   "document": "generated/12345_generated.docx",
     *   "pictures[]": (archivo1.jpg),
     *   "pictures[]": (archivo2.jpg),
     *   "sizes[]": "grande",      ← tamaño de la foto 1
     *   "sizes[]": "mediana",     ← tamaño de la foto 2
     *   "sizes[]": "pequeña"      ← tamaño de la foto 3
     * }
     */
    public function addPictures(Request $request)
    {
        $request->validate([
            'document'   => 'required|string',
            'pictures'   => 'required|array|min:1',
            'pictures.*' => 'required|file|image',
            'sizes'      => 'sometimes|array',
            'sizes.*'    => 'sometimes|string|in:grande,mediana,pequeña',
        ]);

        $documentPath = $request->input('document');
        $sizes = $request->input('sizes', []);

        if (!Storage::exists($documentPath)) {
            return response()->json(['message' => 'Document not found'], 404);
        }

        // 🔹 guardar imágenes temporalmente y obtener rutas absolutas
        $tempRelativePaths = [];
        $absolutePicturePaths = [];

        foreach ($request->file('pictures') as $file) {
            $tempRelative = $file->store('temp/pictures');
            $tempRelativePaths[] = $tempRelative;
            $absolutePicturePaths[] = Storage::path($tempRelative);
        }

        try {
            $this->generateService->addPicturesToDocument($documentPath, $absolutePicturePaths, $sizes);
        } finally {
            // 🔹 limpiar imágenes temporales
            Storage::delete($tempRelativePaths);
        }

        // 🔹 ruta del archivo final generado por el servicio
        $finalAbsolutePath = Storage::path($documentPath) . '_final.docx';

        return response()->download($finalAbsolutePath);
    }
}
