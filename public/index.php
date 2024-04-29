<?php

use Illuminate\Http\Request;
use App\Http\Controllers\BotonExcelController;

define('LARAVEL_START', microtime(true));

// Determine if the application is in maintenance mode...
if (file_exists($maintenance = __DIR__.'/../storage/framework/maintenance.php')) {
    require $maintenance;
}

// Register the Composer autoloader...
require __DIR__.'/../vendor/autoload.php';

// Bootstrap Laravel and handle the request...
$app = require_once __DIR__.'/../bootstrap/app.php';

// Manejar la solicitud
$response = $app->handleRequest(Request::capture());

// Verificar si la respuesta es null
if ($response !== null) {
    // Lógica para exportar Excel
    $botonExcelController = new BotonExcelController();
    $botonExcelController->exportarExcel();

    // Devolver la respuesta
    $response->send();
} else {
    // Manejar el caso en el que la respuesta es null
    echo "";
}