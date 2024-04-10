<?php
 
namespace Database\Seeders;
 
use Illuminate\Database\Seeder;
use App\Models\DiarioCaja;
use Faker\Factory as Faker;
 
class DiarioCajaSeeder extends Seeder
{
    public function run()
    {
        // Creamos una instancia de Faker
        $faker = Faker::create();
 
        // Creamos un solo registro de ejemplo
        DiarioCaja::create([
            'fecha' => $faker->date(),
            'concepto' => $faker->word,
            'documento' => $faker->word,
            'tags' => $faker->word,
            'debe' => $faker->randomFloat(2, 0, 10000),
            'haber' => $faker->randomFloat(2, 0, 10000),
            'saldo' => $faker->randomFloat(2, 0, 10000),
        ]);
    }
}