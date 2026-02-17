<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelController;

Route::prefix('excel')->group(function () {
    // Subir archivo
    Route::post('/upload', [ExcelController::class, 'upload']);
    
    // Obtener hojas disponibles
    Route::post('/hojas', [ExcelController::class, 'obtenerHojas']);
    
    // Obtener encabezados de una hoja
    Route::post('/encabezados', [ExcelController::class, 'obtenerEncabezados']);
    
    // Leer columnas específicas
    Route::post('/leer-columnas', [ExcelController::class, 'leerColumnas']);
});

