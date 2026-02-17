<?php

namespace App\Services;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use Illuminate\Support\Collection;

class ExcelReaderService
{
    public function __construct()
    {
        // Aumentar límite de memoria para archivos grandes
        ini_set('memory_limit', '512M');
        // Aumentar tiempo de ejecución
        ini_set('max_execution_time', '300');
    }

    /**
     * Lee datos de múltiples columnas de un archivo Excel
     *
     * @param string $filePath Ruta del archivo Excel
     * @param array $config Configuración con sheet, columnas y validaciones opcionales
     *                      Ejemplo: [
     *                          'sheet' => 'Hoja1',
     *                          'columnas' => ['PC', 'Nombre'],
     *                          'validaciones' => [
     *                              'PC' => 'numeric' // Detener si PC no es numérico
     *                          ]
     *                      ]
     * @return Collection
     * @throws \Exception
     */
    public function leerColumnas(string $filePath, array $config): Collection
    {
        $sheetName = $config['sheet'];
        $columnasRequeridas = $config['columnas'] ?? [$config['columna1'], $config['columna2']];
        $validaciones = $config['validaciones'] ?? [];

        // Usar ReadFilter para cargar solo lo necesario
        $reader = new Xlsx();
        $reader->setReadDataOnly(true); // No cargar estilos ni fórmulas
        $reader->setLoadSheetsOnly([$sheetName]);
        
        $spreadsheet = $reader->load($filePath);
        
        if (!$spreadsheet->sheetNameExists($sheetName)) {
            throw new \Exception("La hoja '{$sheetName}' no existe en el archivo.");
        }
        
        $sheet = $spreadsheet->getSheetByName($sheetName);
        
        // Encontrar los encabezados de las columnas
        $columnPositions = $this->encontrarEncabezados($sheet, $columnasRequeridas);
        
        // Verificar que se encontraron todas las columnas
        $columnasNoEncontradas = array_diff($columnasRequeridas, array_keys($columnPositions));
        if (!empty($columnasNoEncontradas)) {
            throw new \Exception("No se encontraron las siguientes columnas: " . implode(', ', $columnasNoEncontradas));
        }
        
        // Extraer los datos
        $datos = $this->extraerDatosMultiplesColumnas($sheet, $columnPositions, $columnasRequeridas, $validaciones);
        
        // Liberar memoria
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        
        return collect($datos);
    }

    /**
     * Obtiene los nombres de todas las hojas del archivo
     */
    public function obtenerNombresHojas(string $filePath): array
    {
        $reader = new Xlsx();
        $reader->setReadDataOnly(true);
        
        // Obtener solo los nombres de las hojas sin cargar todo el archivo
        $worksheetNames = $reader->listWorksheetNames($filePath);
        
        return $worksheetNames;
    }

    /**
     * Obtiene los encabezados de una hoja específica
     * Busca en las primeras 30 filas y devuelve encabezados únicos
     */
    public function obtenerEncabezadosHoja(string $filePath, string $sheetName): array
    {
        // Crear un filtro para cargar solo las primeras 30 filas
        $filterSubset = new class(30) implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter {
            private $maxRow;
            
            public function __construct($maxRow) {
                $this->maxRow = $maxRow;
            }
            
            public function readCell($columnAddress, $row, $worksheetName = '') {
                return $row <= $this->maxRow;
            }
        };
        
        $reader = new Xlsx();
        $reader->setReadDataOnly(true);
        $reader->setLoadSheetsOnly([$sheetName]);
        $reader->setReadFilter($filterSubset);
        
        $spreadsheet = $reader->load($filePath);
        
        if (!$spreadsheet->sheetNameExists($sheetName)) {
            throw new \Exception("La hoja '{$sheetName}' no existe en el archivo.");
        }
        
        $sheet = $spreadsheet->getSheetByName($sheetName);
        $encabezados = [];
        $highestColumn = $sheet->getHighestColumn();
        
        // Buscar en las primeras 30 filas
        for ($row = 1; $row <= 30; $row++) {
            $rowHeaders = [];
            $hasMultipleHeaders = false;
            
            for ($col = 'A'; $col <= $highestColumn; $col++) {
                $cellValue = $this->obtenerValorCelda($sheet, $col . $row);
                
                if (!empty($cellValue) && strlen($cellValue) > 1) {
                    $rowHeaders[] = $cellValue;
                }
            }
            
            // Si encontramos una fila con 5 o más encabezados, probablemente sea la fila de encabezados
            if (count($rowHeaders) >= 5) {
                $encabezados = array_merge($encabezados, $rowHeaders);
            }
        }
        
        // Liberar memoria
        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
        
        // Eliminar duplicados manteniendo el orden
        return array_values(array_unique($encabezados));
    }

    /**
     * Encuentra las posiciones de los encabezados en la hoja
     * Busca en las primeras 30 filas
     */
    private function encontrarEncabezados(Worksheet $sheet, array $nombresColumnas): array
    {
        $highestRow = min(30, $sheet->getHighestRow());
        $highestColumn = $sheet->getHighestColumn();
        $columnPositions = [];

        for ($row = 1; $row <= $highestRow; $row++) {
            for ($col = 'A'; $col <= $highestColumn; $col++) {
                $cellValue = $this->obtenerValorCelda($sheet, $col . $row);
                
                if (in_array($cellValue, $nombresColumnas) && !isset($columnPositions[$cellValue])) {
                    $columnPositions[$cellValue] = [
                        'letra' => $col,
                        'fila' => $row
                    ];
                }
            }
            
            // Si ya encontramos todas las columnas, salir
            if (count($columnPositions) === count($nombresColumnas)) {
                break;
            }
        }

        return $columnPositions;
    }

    /**
     * Extrae los datos de múltiples columnas especificadas
     */
    private function extraerDatosMultiplesColumnas(
        Worksheet $sheet,
        array $columnPositions,
        array $nombresColumnas,
        array $validaciones = []
    ): array {
        $datos = [];
        $highestRow = $sheet->getHighestRow();
        $registroNumero = 1;
        
        // Encontrar la fila inicial (la más alta de todas las columnas + 1)
        $filaInicio = 0;
        foreach ($columnPositions as $pos) {
            if ($pos['fila'] > $filaInicio) {
                $filaInicio = $pos['fila'];
            }
        }
        $filaInicio++; // Empezar en la siguiente fila después de los encabezados
        
        // Extraer datos fila por fila
        for ($row = $filaInicio; $row <= $highestRow; $row++) {
            $filaData = [];
            $tieneAlgunValor = false;
            
            // Extraer el valor de cada columna requerida
            foreach ($nombresColumnas as $nombreCol) {
                if (isset($columnPositions[$nombreCol])) {
                    $letra = $columnPositions[$nombreCol]['letra'];
                    $valor = $this->obtenerValorCelda($sheet, $letra . $row);
                    $filaData[$nombreCol] = $valor;
                    
                    if (!empty($valor)) {
                        $tieneAlgunValor = true;
                    }
                }
            }
            
            // Aplicar validaciones si existen
            $debeDetener = false;
            foreach ($validaciones as $columna => $tipoValidacion) {
                if (isset($filaData[$columna])) {
                    $valor = $filaData[$columna];
                    
                    // Si el valor está vacío y hay validación, detener
                    if (empty($valor)) {
                        $debeDetener = true;
                        break;
                    }
                    
                    // Validar según el tipo
                    switch ($tipoValidacion) {
                        case 'numeric':
                            if (!is_numeric($valor)) {
                                $debeDetener = true;
                            }
                            break;
                        case 'not_empty':
                            if (empty(trim($valor))) {
                                $debeDetener = true;
                            }
                            break;
                        case 'date':
                            if (!$this->esFormatoFecha($valor)) {
                                $debeDetener = true;
                            }
                            break;
                    }
                    
                    if ($debeDetener) {
                        break;
                    }
                }
            }
            
            // Si alguna validación falló, detener el procesamiento
            if ($debeDetener) {
                break;
            }
            
            // Solo agregar si al menos una columna tiene valor
            if ($tieneAlgunValor) {
                // Generar row_id único: timestamp + número de registro
                $rowId = time() . sprintf('%04d', $registroNumero);

                // Agregar metadatos del registro
                $filaData['row_id'] = $rowId;
                $filaData['fila'] = $row;
                $filaData['registro_numero'] = $registroNumero;

                $datos[] = $filaData;
                $registroNumero++;
            }
            
            // Liberar memoria cada 1000 filas
            if ($row % 1000 === 0) {
                gc_collect_cycles();
            }
        }

        return $datos;
    }
    
    /**
     * Verifica si un valor tiene formato de fecha
     */
    private function esFormatoFecha(string $valor): bool
    {
        if (empty($valor)) {
            return false;
        }
        
        // Intentar parsear como fecha
        $timestamp = strtotime($valor);
        return $timestamp !== false;
    }

    /**
     * Obtiene el valor de una celda manejando celdas combinadas
     */
    private function obtenerValorCelda(Worksheet $sheet, string $cellAddress): ?string
    {
        try {
            $cell = $sheet->getCell($cellAddress);
            
            // Si la celda está en un rango combinado, obtener el valor de la celda principal
            if ($cell->isInMergeRange()) {
                $mergeRange = $cell->getMergeRange();
                // Obtener la primera celda del rango combinado
                preg_match('/^([A-Z]+\d+)/', $mergeRange, $matches);
                $firstCell = $matches[1] ?? $cellAddress;
                $value = $sheet->getCell($firstCell)->getValue();
            } else {
                $value = $cell->getValue();
            }

            // Convertir a string y limpiar
            if ($value === null || $value === '') {
                return null;
            }
            
            // Si es un objeto DateTime, convertirlo a string
            if ($value instanceof \DateTimeInterface) {
                return $value->format('Y-m-d H:i:s');
            }
            
            return trim((string)$value);
            
        } catch (\Exception $e) {
            return null;
        }
    }
}