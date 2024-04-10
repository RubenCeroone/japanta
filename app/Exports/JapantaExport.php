<?php
 
namespace App\Exports;
 
use App\Models\Japanta;
use App\Models\Japanta1;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
 
class JapantaExport implements FromCollection, WithStyles
{
    protected $japanta;
    protected $japanta1;
 
    public function __construct()
    {
        $this->japanta = Japanta::all();
        $this->japanta->prepend(new Japanta());
    }

    public function __construct1()
    {
        $this->japanta1 = Japanta1::all();
        $this->japanta1->prepend(new Japanta1());
    }
 
    public function collection(): Collection
    {
        return $this->japanta;
        return $this->japanta1;
    }
 
    public function styles(Worksheet $sheet)
    {
        // Establecer el ancho de las columnas
        $sheet->getColumnDimension('A')->setWidth(11.14);
        $sheet->getColumnDimension('B')->setWidth(59.00);
        $sheet->getColumnDimension('C')->setWidth(13.43);
        $sheet->getColumnDimension('D')->setWidth(13.43);
        $sheet->getColumnDimension('E')->setWidth(13.43);
        $sheet->getColumnDimension('F')->setWidth(13.43);
        $sheet->getColumnDimension('G')->setWidth(13.43);

        // Establece el alto de la fila 2 en 24
        $sheet->getRowDimension(2)->setRowHeight(24);

        // Aplicar estilos a la celda A1
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(12);
        $sheet->getStyle('A1')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Escribe el texto en la celda A1
        $sheet->setCellValue('A1', 'Japanta SL');

        // Combina las celdas A2 y B2
        $sheet->mergeCells('A2:B2');
        
        // Aplica estilos al texto en la celda A2
        $sheet->getStyle('A2')->getFont()->setBold(true)->setSize(18);
        $sheet->getStyle('A2')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        
        // Escribe el texto en la celda A2
        $sheet->setCellValue('A2', 'Japanta SL - Libro Mayor 01/09/2023-30/09/2023');

        // Combina las celdas A3 y B3
        $sheet->mergeCells('A3:B3');

        // Aplica estilos al texto en la celda A3
        $sheet->getStyle('A3')->getFont()->setSize(11);
        $sheet->getStyle('A3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Escribe el texto en la celda A3
        $sheet->setCellValue('A3', '01/09/2023 - 30/09/2023');

        // Combina las celdas A4 y B4
        $sheet->mergeCells('A4:B4');

        // Aplica estilos al texto en la celda A4
        $sheet->getStyle('A4')->getFont()->setSize(11);
        $sheet->getStyle('A4')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Escribe el texto en la celda A4
        $sheet->setCellValue('A4', 'añadir desde/hasta de la sección de cuentas contables');

        // Combina las celdas A6 y B6
        $sheet->mergeCells('A6:B6');

        // Aplica estilos al texto en la celda A6
        $sheet->getStyle('A6')->getFont()->setSize(11);
        $sheet->getStyle('A6')->getFont()->setBold(true);
        $sheet->getStyle('A6')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Escribe el texto en la celda A6
        $sheet->setCellValue('A6', '4720000005, Hp, Iva Soportado 5%');

        // Establece el color de fondo de las celdas A7:G7 a negro
        $sheet->getStyle('A7:G7')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('000000');

        // Aplica estilos al texto de las celdas A7:G7
        $sheet->getStyle('A7:G7')->getFont()->setSize(11);
        $sheet->getStyle('A7:G7')->getFont()->setBold(true);
        $sheet->getStyle('A7:G7')->getFont()->setColor(new Color(Color::COLOR_WHITE));
        $sheet->getStyle('A7:G7')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Establece el texto en las celdas A7:G7
        $sheet->setCellValue('A7', 'Fecha');
        $sheet->setCellValue('B7', 'Concepto');
        $sheet->setCellValue('C7', 'Documento');
        $sheet->setCellValue('D7', 'Tags');
        $sheet->setCellValue('E7', 'Debe');
        $sheet->setCellValue('F7', 'Haber');
        $sheet->setCellValue('G7', 'Saldo');

        // Insertar datos de Japanta
        $startRow = 8; // Fila donde empiezan los datos
        $endRow = 14; // Última fila donde se insertarán datos

        $row = $startRow;
        $totalDebe = 0;
        $totalHaber = 0;
        foreach ($this->collection() as $japanta) {
            if ($row > $endRow) {
                break; // Salir del bucle si alcanzamos la última fila
        }
        $sheet->setCellValue('A' . $row, $japanta->fecha);
        $sheet->setCellValue('B' . $row, $japanta->concepto);
        $sheet->setCellValue('C' . $row, $japanta->documento);
        $sheet->setCellValue('D' . $row, $japanta->tags);

        // Formatear números en las columnas debe, haber y saldo
        $sheet->setCellValue('E' . $row, $japanta->debe != 0 ? number_format($japanta->debe, 2) . ' €' : '- €');
        $sheet->setCellValue('F' . $row, $japanta->haber != 0 ? number_format($japanta->haber, 2) . ' €' : '- €');

        // Calcular el saldo sumando el total de debe y restando el total de haber
        $totalDebe += $japanta->debe;
        $totalHaber += $japanta->haber;
        $saldo = $totalDebe - $totalHaber;
        $sheet->setCellValue('G' . $row, $this->formatNumber($saldo) . " €");

        // Aplicar estilos a las celdas de las columnas debe, haber y saldo
        $sheet->getStyle('E' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('F' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('G' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('E' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
        $sheet->getStyle('F' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
        $sheet->getStyle('G' . $row)->getNumberFormat()->setFormatCode('#,##0.00');

        $row++;
    }
        
        // Aplica estilos al texto en la celda B16
        $sheet->getStyle('B16')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle('B16')->getFont()->setSize(11);
        
        // Escribe el texto en la celda B16
        $sheet->setCellValue('B16', 'Total');

        // Calcular la suma de los valores en las celdas E8 hasta E14
        $sumaDebe = '=SUM(E8:E14)';

        // Asignar la fórmula de suma a la celda E16
        $sheet->setCellValue('E16', $sumaDebe . ' €');

        // Aplicar el formato numérico a la celda E16
        $sheet->getStyle('E16')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        
        // Calcular la suma de los valores en las celdas F8 hasta F14
        $sumaHaber = '=SUM(F8:F14)';

        // Asignar la fórmula de suma a la celda F16
        $sheet->setCellValue('F16', $sumaHaber . ' €');

        // Aplicar el formato numérico a la celda F16
        $sheet->getStyle('F16')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        // Calcular la resta de los resultados en las celdas E16 y F16
        $restaSaldo = '=E16-F16';

        // Asignar la fórmula de resta a la celda G16
        $sheet->setCellValue('G16', $restaSaldo . ' €');

        // Aplicar el formato numérico a la celda G16
        $sheet->getStyle('G16')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        // Combinar celdas A18 y B18
        $sheet->mergeCells('A18:B18');

        // Aplica estilos al texto en la celda A18
        $sheet->getStyle('A18')->getFont()->setBold(true)->setSize(11);
        $sheet->getStyle('A18')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Asignar el texto a la celda combinada
        $sheet->setCellValue('A18', '4720000010, Hp, Iva Soportado 10%');

        // Establece el color de fondo de las celdas A19:G19 a negro
        $sheet->getStyle('A19:G19')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('000000');

        // Aplica estilos al texto de las celdas A19:G19
        $sheet->getStyle('A19:G19')->getFont()->setSize(11);
        $sheet->getStyle('A19:G19')->getFont()->setBold(true);
        $sheet->getStyle('A19:G19')->getFont()->setColor(new Color(Color::COLOR_WHITE));
        $sheet->getStyle('A19:G19')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Establece el texto en las celdas A19:G19
        $sheet->setCellValue('A19', 'Fecha');
        $sheet->setCellValue('B19', 'Concepto');
        $sheet->setCellValue('C19', 'Documento');
        $sheet->setCellValue('D19', 'Tags');
        $sheet->setCellValue('E19', 'Debe');
        $sheet->setCellValue('F19', 'Haber');
        $sheet->setCellValue('G19', 'Saldo');

        // Insertar datos de Japanta
        $startRow = 20; // Fila donde empiezan los datos
        $endRow = 30; // Última fila donde se insertarán datos

        $row = $startRow;
        $totalDebe1 = 0;
        $totalHaber1 = 0;
        foreach ($this->collection() as $japanta1) {
            if ($row > $endRow) {
                break; // Salir del bucle si alcanzamos la última fila
        }
        $sheet->setCellValue('A' . $row, $japanta1->fecha1);
        $sheet->setCellValue('B' . $row, $japanta1->concepto1);
        $sheet->setCellValue('C' . $row, $japanta1->documento1);
        $sheet->setCellValue('D' . $row, $japanta1->tags1);

        // Formatear números en las columnas debe, haber y saldo
        $sheet->setCellValue('E' . $row, $japanta1->debe1 != 0 ? number_format($japanta1->debe1, 2) . ' €' : '- €');
        $sheet->setCellValue('F' . $row, $japanta1->haber1 != 0 ? number_format($japanta1->haber1, 2) . ' €' : '- €');

        // Calcular el saldo sumando el total de debe y restando el total de haber
        $totalDebe1 += $japanta1->debe1;
        $totalHaber1 += $japanta1->haber1;
        $saldo1 = $totalDebe1 - $totalHaber1;
        $sheet->setCellValue('G' . $row, $this->formatNumber($saldo1) . " €");

        // Aplicar estilos a las celdas de las columnas debe, haber y saldo
        $sheet->getStyle('E' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('F' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('G' . $row)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('E' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
        $sheet->getStyle('F' . $row)->getNumberFormat()->setFormatCode('#,##0.00');
        $sheet->getStyle('G' . $row)->getNumberFormat()->setFormatCode('#,##0.00');

        $row++;
    }

    // Aplica estilos al texto en la celda B32
    $sheet->getStyle('B32')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
    $sheet->getStyle('B32')->getFont()->setSize(11);
    
    // Escribe el texto en la celda B32
    $sheet->setCellValue('B32', 'Total');

    // Calcular la suma de los valores en las celdas E20 hasta E30
    $sumaDebe1 = '=SUM(E20:E30)';

    // Asignar la fórmula de suma a la celda E32
    $sheet->setCellValue('E32', $sumaDebe1 . ' €');

    // Aplicar el formato numérico a la celda E32
    $sheet->getStyle('E32')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
    
    // Calcular la suma de los valores en las celdas F20 hasta F30
    $sumaHaber1 = '=SUM(F20:F30)';

    // Asignar la fórmula de suma a la celda F32
    $sheet->setCellValue('F32', $sumaHaber1 . ' €');

    // Aplicar el formato numérico a la celda F32
    $sheet->getStyle('F32')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

    // Calcular la resta de los resultados en las celdas E32 y F32
    $restaSaldo = '=E32-F32';

    // Asignar la fórmula de resta a la celda G32
    $sheet->setCellValue('G32', $restaSaldo1 . ' €');

    // Aplicar el formato numérico a la celda G32
    $sheet->getStyle('G32')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
    }
}