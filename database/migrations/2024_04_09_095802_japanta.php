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
            Schema::create('japanta', function (Blueprint $table) {
                $table->id();
                $table->date('fecha');
                $table->string('concepto');
                $table->string('documento');
                $table->string('tags');
                $table->decimal('debe', 10, 2);
                $table->decimal('haber', 10, 2);
                $table->decimal('saldo', 10, 2);
                $table->timestamps();
        });
    }
 
    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('japanta');
    }
};