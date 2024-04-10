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
            Schema::create('japanta1', function (Blueprint $table) {
                $table->id();
                $table->date('fecha1');
                $table->string('concepto1');
                $table->string('documento1');
                $table->string('tags1');
                $table->decimal('debe1', 10, 2);
                $table->decimal('haber1', 10, 2);
                $table->decimal('saldo1', 10, 2);
                $table->timestamps();
        });
    }
 
    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('japanta1');
    }
};