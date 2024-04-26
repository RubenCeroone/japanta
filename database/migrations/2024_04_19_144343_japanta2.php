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
            Schema::create('japanta2', function (Blueprint $table) {
                $table->id();
                $table->date('fecha2');
                $table->string('nombre');
                $table->string('grupo');
                $table->string('saldoinicial');
                $table->decimal('debe2', 10, 2);
                $table->decimal('haber2', 10, 2);
                $table->decimal('saldofinal', 10, 2);
                $table->timestamps();
        });
    }
 
    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('japanta2');
    }
};