<?php

use Illuminate\Foundation\Application;
use Illuminate\Support\Facades\Route;
use App\Http\Controllers\BotonExcelController;
use Inertia\Inertia;
 
 
 
Route::get('/BotonExcel',[BotonExcelController::class,'invoke'])->name('botonexcel');
Route::get('/exportar-excel', [BotonExcelController::class, 'exportarExcel'])->name('exportar_excel');

/*
Route::get('/', function () {
    return Inertia::render('Welcome', [
        'canLogin' => Route::has('login'),
        'canRegister' => Route::has('register'),
        'laravelVersion' => Application::VERSION,
        'phpVersion' => PHP_VERSION,
    ]);
});

Route::middleware([
    'auth:sanctum',
    config('jetstream.auth_session'),
    'verified',
])->group(function () {
    Route::get('/dashboard', function () {
        return Inertia::render('Dashboard');
    })->name('dashboard');
});
*/