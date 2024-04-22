<?php

namespace App\Http\Controllers;

use App\Http\Controllers\Controller;
use App\Models\Japanta2;
use Maatwebsite\Excel\Facades\Excel;
use App\Exports\JapantaExport;

class BotonExcelController extends Controller
{
    public function invoke()
    {
        return view('botonexcel');
    }

    public function exportarExcel()
    {
        $datos = Japanta2::all();
        return Excel::download(new JapantaExport2($datos), 'Japanta2.xlsx');
    }
}
