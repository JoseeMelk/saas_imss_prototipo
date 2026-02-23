<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExcelController;
use App\Http\Controllers\DocumentTemplateController;
use App\Http\Controllers\DocumentGenerateController;

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

Route::prefix('templates')->group(function () {
    // Subir plantilla
    Route::post('/upload', [DocumentTemplateController::class, 'upload']);
    
    // Listar plantillas
    Route::get('/', [DocumentTemplateController::class, 'index']);
});

Route::post('/generate', [DocumentGenerateController::class, 'generate']);
Route::post('/generate/add-pictures', [DocumentGenerateController::class, 'addPictures']);
