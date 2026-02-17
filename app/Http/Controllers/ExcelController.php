<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Services\ExcelBatchService;
use App\Services\ExcelReaderService;
use Illuminate\Http\JsonResponse;
use App\Models\ExcelData;

class ExcelController extends Controller
{
    public function upload(Request $request, ExcelBatchService $batchService)
    {
        $request->validate([
            'file' => 'required|file|mimes:xlsx,xls,xlsm'
        ]);

        $batch = $batchService->create($request->file('file'));

        return response()->json([
            'batch_id' => $batch->id,
            'file_name' => $batch->original_name
        ]);
    }

    /**
     * Lee columnas específicas de un archivo Excel
     * Soporta 2 columnas (compatibilidad) o múltiples columnas
     * Soporta validaciones para detener lectura cuando una columna no cumple criterios
     */
    public function leerColumnas(
        Request $request,
        ExcelBatchService $batchService,
        ExcelReaderService $readerService
    ): JsonResponse {
        $request->validate([
            'batch_id' => 'required|integer|exists:excel_batches,id',
            'sheet' => 'required|string',
            'columnas' => 'array|min:1', // Array de columnas
            'columna1' => 'string', // Compatibilidad con versión anterior
            'columna2' => 'string', // Compatibilidad con versión anterior
            'validaciones' => 'array', // Validaciones opcionales
        ]);

        try {
            // Obtener la ruta del archivo desde el batch
            $filePath = $batchService->getPath($request->input('batch_id'));

            // Determinar qué columnas se van a leer
            if ($request->has('columnas')) {
                $columnas = $request->input('columnas');
            } else {
                // Compatibilidad con versión anterior (columna1, columna2)
                $columnas = array_filter([
                    $request->input('columna1'),
                    $request->input('columna2')
                ]);
            }

            if (empty($columnas)) {
                return response()->json([
                    'success' => false,
                    'mensaje' => 'Debe especificar al menos una columna'
                ], 400);
            }

            // Configuración
            $config = [
                'sheet' => $request->input('sheet'),
                'columnas' => $columnas
            ];

            // Agregar validaciones si existen
            if ($request->has('validaciones')) {
                $config['validaciones'] = $request->input('validaciones');
            }

            // Leer datos
            $datos = $readerService->leerColumnas($filePath, $config);

            ExcelData::create([
                'batch_id' => $request->input('batch_id'),
                'sheet' => $config['sheet'],
                'columns' => implode(',', $columnas),
                'datas' => json_encode($datos)
            ]);
            return response()->json([
                'success' => true,
                'total' => $datos->count(),
                'columnas' => $columnas,
                'validaciones' => $config['validaciones'] ?? null,
                'datos' => $datos
            ]);

        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'mensaje' => $e->getMessage()
            ], 400);
        }
    }

    /**
     * Obtener las hojas disponibles en un archivo Excel
     */
    public function obtenerHojas(
        Request $request,
        ExcelBatchService $batchService,
        ExcelReaderService $readerService
    ): JsonResponse {
        $request->validate([
            'batch_id' => 'required|integer|exists:excel_batches,id',
        ]);

        try {
            $filePath = $batchService->getPath($request->input('batch_id'));
            $hojas = $readerService->obtenerNombresHojas($filePath);

            return response()->json([
                'success' => true,
                'hojas' => $hojas
            ]);

        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'mensaje' => $e->getMessage()
            ], 400);
        }
    }

    /**
     * Obtener los encabezados de columnas de una hoja específica
     */
    public function obtenerEncabezados(
        Request $request,
        ExcelBatchService $batchService,
        ExcelReaderService $readerService
    ): JsonResponse {
        $request->validate([
            'batch_id' => 'required|integer|exists:excel_batches,id',
            'sheet' => 'required|string',
        ]);

        try {
            $filePath = $batchService->getPath($request->input('batch_id'));
            $encabezados = $readerService->obtenerEncabezadosHoja(
                $filePath,
                $request->input('sheet')
            );

            return response()->json([
                'success' => true,
                'total' => count($encabezados),
                'encabezados' => $encabezados
            ]);

        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'mensaje' => $e->getMessage()
            ], 400);
        }
    }
}