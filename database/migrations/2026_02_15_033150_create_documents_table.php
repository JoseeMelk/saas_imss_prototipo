<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::create('documents', function (Blueprint $table) {
            $table->id();
            $table->string('name');
            $table->string('path');
            $table->string('pc');//Punto de control
            $table->string('ic');//Incumplimiento
            $table->string('ac');//Accion correctiva
            $table->boolean('has_images')->default(false);
            $table->boolean('is_completed')->default(false);//Avance
            $table->string('row_id');//id de la fila procesada del json
            $table->foreignId('batch_id')->constrained('excel_batches')->onDelete('cascade');
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('documents');
    }
};
