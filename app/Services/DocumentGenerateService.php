<?php

namespace App\Services;

use App\Models\Document;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\SimpleType\Jc;
use Illuminate\Support\Facades\Storage;
use Illuminate\Support\Facades\Log;

class DocumentGenerateService
{
    /**
     * Valida que los datos requeridos estén presentes
     */
    private function validateData(array $data): void
    {
        $required = ['ic', 'ac', 'row_id', 'batch_id'];

        foreach ($required as $field) {
            if (empty($data[$field])) {
                throw new \InvalidArgumentException(
                    "El campo '{$field}' es requerido"
                );
            }
        }
    }

    /**
     * Genera un documento Word con formato profesional
     * 
     * @param array $data Debe contener: ic, ac, row_id, batch_id, pc (opcional)
     * @return Document
     */
    public function generateDocument(array $data): Document
    {
        // Validar datos requeridos
        $this->validateData($data);

        // Generar nombre único del documento
        $pc = $data['pc'] ?? $data['row_id'];
        $filename = "PC_{$pc}_{$data['row_id']}.docx";
        $relativePath = "documents/{$data['batch_id']}/{$filename}";
        $fullPath = storage_path("app/{$relativePath}");

        // Crear directorio si no existe
        $directory = dirname($fullPath);
        if (!file_exists($directory)) {
            mkdir($directory, 0755, true);
        }

        try {
            // Generar documento con PhpWord
            $this->generateWithPhpWord($data, $fullPath);

            // Guardar en base de datos
            $document = Document::create([
                'name' => $filename,
                'path' => $relativePath,
                'ic' => $data['ic'],
                'ac' => $data['ac'],
                'pc' => $data['pc'] ?? null,
                'is_completed' => false,
                'has_images' => false,
                'row_id' => $data['row_id'],
                'batch_id' => $data['batch_id'],
            ]);

            return $document;
        } catch (\Exception $e) {
            Log::error("Error generando documento", [
                'row_id' => $data['row_id'],
                'batch_id' => $data['batch_id'],
                'error' => $e->getMessage()
            ]);
            throw $e;
        }
    }

    /**
     * Genera documento usando PhpWord con formato profesional
     */
    private function generateWithPhpWord(array $data, string $outputPath): void
    {
        $phpWord = new PhpWord();

        // Configurar estilos por defecto
        $phpWord->setDefaultFontName('Arial');
        $phpWord->setDefaultFontSize(11);

        // Configurar estilos personalizados
        $phpWord->addParagraphStyle('Title', [
            'alignment' => Jc::CENTER,
            'spaceAfter' => 400,
        ]);

        $phpWord->addParagraphStyle('Heading', [
            'spaceAfter' => 100,
        ]);

        $phpWord->addParagraphStyle('Content', [
            'alignment' => Jc::BOTH,
            'spaceAfter' => 300,
        ]);

        // Crear sección con márgenes
        $section = $phpWord->addSection([
            'marginTop' => 1440,    // 1 pulgada = 1440 twips
            'marginRight' => 1440,
            'marginBottom' => 1440,
            'marginLeft' => 1440
        ]);

        // ===== TÍTULO =====
        $section->addText(
            'SEGUIMIENTO DE SUPERVISIÓN',
            [
                'bold' => true,
                'size' => 16,
                'color' => '000000'
            ],
            'Title'
        );

        // Línea separadora
        $section->addTextBreak(1);

        // ===== PUNTO DE CONTROL (si existe) =====
        if (!empty($data['pc'])) {
            $section->addText(
                "PUNTO DE CONTROL: {$data['pc']}",
                [
                    'bold' => true,
                    'size' => 14,
                    'color' => '000000'
                ],
                [
                    'spaceAfter' => 200,
                    'spaceBefore' => 100
                ]
            );

            // Línea separadora
            $section->addLine([
                'weight' => 1,
                'width' => 450,
                'height' => 0,
                'color' => 'CCCCCC'
            ]);

            $section->addTextBreak(1);
        }

        // ===== INCUMPLIMIENTO =====
        $section->addText(
            'INCUMPLIMIENTO:',
            [
                'bold' => true,
                'size' => 12,
                'color' => '4472C4', // Azul
                'underline' => 'single'
            ],
            'Heading'
        );

        $section->addText(
            $data['ic'],
            [
                'size' => 11,
                'color' => '000000'
            ],
            'Content'
        );

        $section->addTextBreak(1);

        // ===== ACCIÓN CORRECTIVA =====
        $section->addText(
            'ACCIÓN CORRECTIVA:',
            [
                'bold' => true,
                'size' => 12,
                'color' => '70AD47', // Verde
                'underline' => 'single'
            ],
            'Heading'
        );

        $section->addText(
            $data['ac'],
            [
                'size' => 11,
                'color' => '000000'
            ],
            'Content'
        );

        $section->addTextBreak(2);

        // ===== EVIDENCIA FOTOGRÁFICA =====
        $section->addText(
            'EVIDENCIA FOTOGRÁFICA',
            [
                'bold' => true,
                'size' => 12,
                'color' => '000000'
            ],
            [
                'alignment' => Jc::CENTER,
                'spaceAfter' => 200,
                'spaceBefore' => 400,
                'borderTop' => [
                    'borderStyle' => 'single',
                    'borderSize' => 6,
                    'borderColor' => '000000'
                ],
                'borderBottom' => [
                    'borderStyle' => 'single',
                    'borderSize' => 6,
                    'borderColor' => '000000'
                ]
            ]
        );

        $section->addText(
            '{{IMAGENES}}',
            [
                'size' => 10,
                'color' => 'CCCCCC',
                'italic' => true
            ],
            [
                'alignment' => Jc::CENTER
            ]
        );

        $section->addTextBreak(1);

        // ===== SECCIÓN DE OBSERVACIONES =====
        $section->addText(
            'OBSERVACIONES:',
            [
                'bold' => true,
                'size' => 11,
                'color' => '000000'
            ],
            [
                'spaceAfter' => 100,
                'spaceBefore' => 200
            ]
        );

        // Agregar líneas para escribir observaciones
        for ($i = 0; $i < 5; $i++) {
            $section->addText(
                str_repeat('_', 80),
                ['size' => 11, 'color' => 'CCCCCC']
            );
        }

        // Guardar documento
        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save($outputPath);

        // Verificar que se creó correctamente
        if (!file_exists($outputPath)) {
            throw new \Exception("El archivo no se generó correctamente");
        }
    }

    /**
     * Inserta imágenes en un documento existente
     * 
     * @param Document $document
     * @param array $imagePaths Array de rutas absolutas de imágenes
     * @return bool
     */
    public function insertImages(Document $document, array $imagePaths): bool
    {
        if (empty($imagePaths)) {
            return false;
        }

        $fullPath = storage_path('app/' . $document->path);

        if (!file_exists($fullPath)) {
            throw new \Exception("Documento no encontrado: {$fullPath}");
        }

        try {
            // Desempaquetar documento
            $unpackedPath = storage_path('app/temp/unpacked_' . uniqid());
            $this->unpackDocument($fullPath, $unpackedPath);

            // Generar XML de imágenes
            $imageXml = $this->generateImageXml($imagePaths, $unpackedPath);

            // Reemplazar marcador en document.xml
            $documentXmlPath = "{$unpackedPath}/word/document.xml";
            $content = file_get_contents($documentXmlPath);

            // Remover el marcador de texto
            $content = str_replace('{{IMAGENES}}', '', $content);

            // Buscar el párrafo que contenía el marcador y agregar las imágenes
            $pattern = '/<w:t[^>]*>\s*<\/w:t>/';
            $content = preg_replace($pattern, '<w:t></w:t>' . $imageXml, $content, 1);

            file_put_contents($documentXmlPath, $content);

            // Re-empaquetar documento
            $this->packDocument($unpackedPath, $fullPath);

            // Limpiar
            $this->deleteDirectory($unpackedPath);

            // Actualizar registro en DB
            $document->update(['has_images' => true]);

            return true;
        } catch (\Exception $e) {
            Log::error("Error insertando imágenes", [
                'document_id' => $document->id,
                'error' => $e->getMessage()
            ]);
            throw $e;
        }
    }

    /**
     * Desempaqueta un documento .docx
     */
    private function unpackDocument(string $docPath, string $outputPath): void
    {
        $zip = new \ZipArchive();
        if ($zip->open($docPath) === true) {
            $zip->extractTo($outputPath);
            $zip->close();
        } else {
            throw new \Exception("Error desempaquetando documento");
        }
    }

    /**
     * Re-empaqueta un documento .docx
     */
    private function packDocument(string $unpackedPath, string $outputPath): void
    {
        $zip = new \ZipArchive();
        if ($zip->open($outputPath, \ZipArchive::CREATE | \ZipArchive::OVERWRITE) === true) {
            $files = new \RecursiveIteratorIterator(
                new \RecursiveDirectoryIterator($unpackedPath),
                \RecursiveIteratorIterator::LEAVES_ONLY
            );

            foreach ($files as $file) {
                if (!$file->isDir()) {
                    $filePath = $file->getRealPath();
                    $relativePath = substr($filePath, strlen($unpackedPath) + 1);
                    $zip->addFile($filePath, str_replace('\\', '/', $relativePath));
                }
            }

            $zip->close();
        } else {
            throw new \Exception("Error empaquetando documento");
        }
    }

    /**
     * Genera XML para insertar imágenes
     */
    private function generateImageXml(array $imagePaths, string $unpackedPath): string
    {
        $imageXml = '';
        $mediaPath = "{$unpackedPath}/word/media";

        if (!file_exists($mediaPath)) {
            mkdir($mediaPath, 0755, true);
        }

        foreach ($imagePaths as $index => $imagePath) {
            if (!file_exists($imagePath)) {
                continue;
            }

            $extension = strtolower(pathinfo($imagePath, PATHINFO_EXTENSION));
            $rId = "rId" . (100 + $index);
            $imageName = "image" . ($index + 1) . "." . $extension;
            $destPath = "{$mediaPath}/{$imageName}";

            copy($imagePath, $destPath);
            $this->addImageRelationship($unpackedPath, $rId, $imageName);

            // Obtener dimensiones de la imagen
            list($width, $height) = getimagesize($imagePath);
            $ratio = $width / $height;

            // Tamaño objetivo: 4 pulgadas de ancho
            $targetWidth = 914400 * 4; // EMUs (English Metric Units)
            $targetHeight = (int)($targetWidth / $ratio);

            $imageXml .= $this->getImageXml($rId, $index, $targetWidth, $targetHeight);
        }

        return $imageXml;
    }

    /**
     * Agrega relación de imagen al documento
     */
    private function addImageRelationship(string $unpackedPath, string $rId, string $imageName): void
    {
        $relsPath = "{$unpackedPath}/word/_rels/document.xml.rels";

        if (!file_exists($relsPath)) {
            // Crear archivo de relaciones si no existe
            $relsDir = dirname($relsPath);
            if (!file_exists($relsDir)) {
                mkdir($relsDir, 0755, true);
            }

            $content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' .
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' .
                '</Relationships>';
            file_put_contents($relsPath, $content);
        } else {
            $content = file_get_contents($relsPath);
        }

        $relationship = '<Relationship Id="' . $rId . '" ' .
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' .
            'Target="media/' . $imageName . '"/>';

        $content = str_replace('</Relationships>', $relationship . '</Relationships>', $content);
        file_put_contents($relsPath, $content);
    }

    /**
     * Genera XML para una imagen individual
     */
    private function getImageXml(string $rId, int $index, int $width, int $height): string
    {
        return <<<XML
</w:t></w:r></w:p>
<w:p>
  <w:r>
    <w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="{$width}" cy="{$height}"/>
        <wp:docPr id="{$index}" name="Imagen {$index}"/>
        <wp:cNvGraphicFramePr>
          <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
        </wp:cNvGraphicFramePr>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="{$index}" name="Imagen {$index}"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="{$rId}"/>
                <a:stretch>
                  <a:fillRect/>
                </a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="{$width}" cy="{$height}"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                  <a:avLst/>
                </a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  </w:r>
</w:p>
<w:p><w:r><w:t>
XML;
    }

    /**
     * Elimina un directorio recursivamente
     */
    private function deleteDirectory(string $path): void
    {
        if (!file_exists($path)) {
            return;
        }

        $files = new \RecursiveIteratorIterator(
            new \RecursiveDirectoryIterator($path, \RecursiveDirectoryIterator::SKIP_DOTS),
            \RecursiveIteratorIterator::CHILD_FIRST
        );

        foreach ($files as $fileinfo) {
            $todo = ($fileinfo->isDir() ? 'rmdir' : 'unlink');
            $todo($fileinfo->getRealPath());
        }

        rmdir($path);
    }
}
