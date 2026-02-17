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
        Schema::create('excel_data', function (Blueprint $table) {
            $table->id();
            $table->string('sheet');
            $table->text('columns'); // Guardar columnas leídas como string separado por
            $table->mediumText('datas');
            $table->foreignId('batch_id')->constrained('excel_batches')->onDelete('cascade');
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('excel_data');
    }
};
