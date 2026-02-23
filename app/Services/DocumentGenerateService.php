<?php

namespace App\Services;

use Illuminate\Support\Facades\Log;
use PhpOffice\PhpWord\TemplateProcessor;
use Illuminate\Support\Facades\Storage;

class DocumentGenerateService
{
    /**
     * Tamaños disponibles para las imágenes (ancho × alto en px a 96 DPI).
     * grande  → 16 cm × 12 cm — 1 foto por fila, foto_der vacío
     * mediana →  6 cm × 4.5 cm — 2 fotos por fila
     * pequeña →  4 cm × 3 cm   — 2 fotos por fila
     *
     * La plantilla siempre tiene el mismo bloque:
     *   ${evidencias} / tabla 2 cols: ${foto_izq} | ${foto_der} / ${/evidencias}
     */
    private const SIZES = [
        'grande'  => ['width' => 605, 'height' => 454, 'perRow' => 1],
        'mediana' => ['width' => 320, 'height' => 245, 'perRow' => 2], // 9 cm × 7 cm
        'pequeña' => ['width' => 151, 'height' => 113, 'perRow' => 2],
    ];

    public function generateDocument(string $templatePath, array $data, string $relativeOutputPath): void
    {
        try {
            // 🔹 crear archivo vacío (crea carpeta automáticamente)
            Storage::put($relativeOutputPath, '');

            // 🔹 obtener ruta real
            $absolutePath = Storage::path($relativeOutputPath);

            $templateProcessor = new TemplateProcessor($templatePath);

            foreach ($data as $key => $value) {
                $templateProcessor->setValue($key, htmlspecialchars($value));
            }

            $templateProcessor->saveAs($absolutePath);
        } catch (\Throwable $e) {
            Log::error("Error al generar documento: " . $e->getMessage());
            throw $e;
        }
    }

    /**
     * Agrupa imágenes en filas según su tamaño individual:
     * - grande  → siempre sola (1 por fila)
     * - mediana/pequeña → se empareja con la siguiente si también es no-grande
     */
    private function buildRows(array $pictures, array $sizes): array
    {
        $rows = [];
        $i = 0;
        $total = count($pictures);

        while ($i < $total) {
            $size = $sizes[$i] ?? 'grande';

            if ($size === 'grande') {
                $rows[] = [
                    ['path' => $pictures[$i], 'size' => $size],
                ];
                $i++;
            } else {
                $nextSize = $sizes[$i + 1] ?? null;
                if ($nextSize !== null && $nextSize !== 'grande') {
                    $rows[] = [
                        ['path' => $pictures[$i],     'size' => $size],
                        ['path' => $pictures[$i + 1], 'size' => $nextSize],
                    ];
                    $i += 2;
                } else {
                    $rows[] = [
                        ['path' => $pictures[$i], 'size' => $size],
                    ];
                    $i++;
                }
            }
        }

        return $rows;
    }

    public function addPicturesToDocument(string $documentPath, array $pictures, array $sizes = []): void
    {
        try {
            $absolutePath = Storage::path($documentPath);
            $templateProcessor = new TemplateProcessor($absolutePath);

            if (empty($pictures)) {
                Log::warning("No se proporcionaron imágenes.");
                return;
            }

            // Normalizar sizes: completar con 'grande' si faltan
            $sizes = array_values($sizes);
            foreach ($pictures as $i => $_) {
                if (!isset($sizes[$i])) {
                    $sizes[$i] = 'grande';
                }
            }

            $rows = $this->buildRows(array_values($pictures), $sizes);

            // Un solo bloque: ${evidencias} / ${foto_izq} | ${foto_der} / ${/evidencias}
            $templateProcessor->cloneBlock('evidencias', count($rows), true, true);

            foreach ($rows as $index => $pair) {
                $n = $index + 1;

                $leftDim = self::SIZES[$pair[0]['size']] ?? self::SIZES['grande'];
                $templateProcessor->setImageValue("foto_izq#$n", [
                    'path'   => $pair[0]['path'],
                    'width'  => $leftDim['width'],
                    'height' => $leftDim['height'],
                ]);

                if (isset($pair[1])) {
                    $rightDim = self::SIZES[$pair[1]['size']] ?? self::SIZES['grande'];
                    $templateProcessor->setImageValue("foto_der#$n", [
                        'path'   => $pair[1]['path'],
                        'width'  => $rightDim['width'],
                        'height' => $rightDim['height'],
                    ]);
                } else {
                    $templateProcessor->setValue("foto_der#$n", '');
                }
            }

            $templateProcessor->saveAs($absolutePath . '_final.docx');
        } catch (\Throwable $e) {
            Log::error("Error al agregar imágenes: " . $e->getMessage());
            throw $e;
        }
    }
}
