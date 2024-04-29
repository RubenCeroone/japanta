<?php

namespace App\Exports;

use App\Models\Japanta;
use App\Models\Japanta1;
use Illuminate\Support\Collection;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithStyles;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use Carbon\Carbon;

class JapantaExport implements FromCollection, WithStyles
{
    protected $japanta;
    protected $japanta1;

    public function __construct()
    {
        $this->japanta = Japanta::all();
        $this->japanta->prepend(new Japanta());
        $this->japanta1 = Japanta1::all();
        $this->japanta1->prepend(new Japanta1());
    }

    public function collection(): Collection
    {
        return $this->japanta->merge($this->japanta1);
    }

    public function formatNumber($number)
    {
        return number_format($number, 2); // Formatear a dos decimales
    }

    public function clearRow1(Worksheet $sheet)
    {
        // Borra los datos de la línea 1 desde la columna C hasta la columna V
        for ($column = 'B'; $column <= 'H'; $column++) {
            $sheet->setCellValue($column . '1', '');
        }
    }

    public function clearRow2(Worksheet $sheet)
    {
        // Borra los datos desde la celda F2 hasta la celda H4
        for ($row = 2; $row <= 6; $row++) {
            for ($column = 'F'; $column <= 'H'; $column++) {
                $sheet->setCellValue($column . $row, '');
            }
        }
    }

    public function clearRow3(Worksheet $sheet)
    {
        // Borra los datos desde la celda H1 hasta la celda H2000
        for ($row = 1; $row <= 2000; $row++) {
            $sheet->setCellValue('H' . $row, '');
        }
    }

    public function clearRow4(Worksheet $sheet)
    {
        // Borra los datos desde la celda C2 hasta la celda D6
        for ($row = 2; $row <= 6; $row++) {
            for ($column = 'C'; $column <= 'D'; $column++) {
                $sheet->setCellValue($column . $row, '');
            }
        }
    }

    public function addBottomBorder(Worksheet $sheet, $startRow, $endRow, $startColumn, $endColumn)
    {
        // Asumimos que la última fila con datos es la inicial
        $lastRowWithData = $startRow; // Comenzamos desde la fila inicial

        // Iterar hacia adelante desde la fila inicial hasta el fin del rango deseado
        for ($row = $startRow; $row <= $endRow; $row++) {
            // Verificar si al menos una celda en el rango tiene contenido
            $isEmpty = true; // Asignamos verdadero inicialmente
            for ($col = $startColumn; ord($col) <= ord($endColumn); $col++) {
                if (!empty(trim($sheet->getCell($col . $row)->getValue()))) {
                    $isEmpty = false; // Encontramos datos, marcar como no vacío
                    break; // Salir del bucle interno si encuentra datos
                }
            }
            if (!$isEmpty) {
                $lastRowWithData = $row; // Actualizar la última fila con datos
            }
        }

        // Agregar "Total:" en negrita a la izquierda de una fila debajo del último dato
        $totalRow = $lastRowWithData + 1;
        // Verificar si la fila total está dentro del rango, ajustar si es necesario
        if ($totalRow > $endRow) {
            $totalRow = $endRow + 1;
        }
        $totalColumn = 'B'; // Columna B
        $sheet->getStyle($totalColumn . $totalRow)->getFont()->setBold(true);
        $sheet->getStyle($totalColumn . $totalRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue($totalColumn . $totalRow, "Total:");
    }
/*
    public function insertTotal($sheet, $startRow, $endRow, $totalDebe, $totalHaber)
    {
        // Calcular el saldo final
        $saldoFinal = $totalDebe - $totalHaber;

        // Calcular la fila del total (2 filas debajo del endRow)
        $totalRow = $endRow + 2;

        // Asignar los valores de total y saldo a las celdas correspondientes
        $sheet->setCellValue('B' . $totalRow, 'Total:');
        $sheet->setCellValue('E' . $totalRow, $totalDebe . ' €');
        $sheet->setCellValue('F' . $totalRow, $totalHaber . ' €');
        $sheet->setCellValue('G' . $totalRow, $saldoFinal . ' €');

        // Aplicar estilos al texto en la celda del total
        $sheet->getStyle('B' . $totalRow)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle('B' . $totalRow)->getFont()->setSize(11);

        // Aplicar el formato numérico a las celdas de total
        $sheet->getStyle('E' . $totalRow)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('F' . $totalRow)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
        $sheet->getStyle('G' . $totalRow)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        // Calcular la suma de los valores en las celdas E8 hasta E(endRow)
        $sumaDebe = '=SUM(E' . $startRow . ':E' . $endRow . ')';
        $sheet->setCellValue('E' . ($totalRow + 1), $sumaDebe . ' €');
        $sheet->getStyle('E' . ($totalRow + 1))->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        // Calcular la suma de los valores en las celdas F8 hasta F(endRow)
        $sumaHaber = '=SUM(F' . $startRow . ':F' . $endRow . ')';
        $sheet->setCellValue('F' . ($totalRow + 1), $sumaHaber . ' €');
        $sheet->getStyle('F' . ($totalRow + 1))->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        // Calcular la resta de los resultados en las celdas E(totalRow+1) y F(totalRow+1)
        $restaSaldo = '=E' . ($totalRow + 1) . '-F' . ($totalRow + 1);
        $sheet->setCellValue('G' . ($totalRow + 1), $restaSaldo . ' €');
        $sheet->getStyle('G' . ($totalRow + 1))->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        return $totalRow; // Devuelve la fila del total
    }

    public function insertTotalSecond($sheet, $startRow, $endRow, $totalDebe, $totalHaber)
    {
        // Obtener el número de fila donde termina la inserción de datos en la función anterior
        $previousEndRow = $this->insertTotal($sheet, $startRow, $endRow, $totalDebe, $totalHaber);

        // Calcular el número de fila de inicio para la nueva inserción
        $startRowSecond = $previousEndRow + 2;

        // Combinar celdas A y B para la nueva sección
        $sheet->mergeCells('A' . $startRowSecond . ':B' . $startRowSecond);

        // Aplica estilos al texto en la celda A$startRowSecond
        $sheet->getStyle('A' . $startRowSecond)->getFont()->setBold(true)->setSize(11);
        $sheet->getStyle('A' . $startRowSecond)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Asignar el texto a la celda combinada
        $sheet->setCellValue('A' . $startRowSecond, '4720000010, Hp, Iva Soportado 10%');

        // Establece el color de fondo de las celdas A$startRowSecond:G$startRowSecond a negro
        $sheet->getStyle('A' . $startRowSecond . ':G' . $startRowSecond)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('000000');

        // Aplica estilos al texto de las celdas A$startRowSecond:G$startRowSecond
        $sheet->getStyle('A' . $startRowSecond . ':G' . $startRowSecond)->getFont()->setSize(11);
        $sheet->getStyle('A' . $startRowSecond . ':G' . $startRowSecond)->getFont()->setBold(true);
        $sheet->getStyle('A' . $startRowSecond . ':G' . $startRowSecond)->getFont()->setColor(new Color(Color::COLOR_WHITE));
        $sheet->getStyle('A' . $startRowSecond . ':G' . $startRowSecond)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);

        // Establece el texto en las celdas A$startRowSecond:G$startRowSecond
        $sheet->setCellValue('A' . $startRowSecond, 'Fecha');
        $sheet->setCellValue('B' . $startRowSecond, 'Concepto');
        $sheet->setCellValue('C' . $startRowSecond, 'Documento');
        $sheet->setCellValue('D' . $startRowSecond, 'Tags');
        $sheet->setCellValue('E' . $startRowSecond, 'Debe');
        $sheet->setCellValue('F' . $startRowSecond, 'Haber');
        $sheet->setCellValue('G' . $startRowSecond, 'Saldo');

        // Insertar datos de Japanta1
        $totalDebeSecond = 0;
        $totalHaberSecond = 0;
        $endRowSecond = $endRow + $startRowSecond; // Se calcula el número total de filas

        for ($row = $startRowSecond + 1; $row <= $endRowSecond; $row++) {
            // Agregar aquí tu lógica para insertar datos de Japanta1
            // Asegúrate de actualizar $totalDebeSecond y $totalHaberSecond en cada iteración
        }

        // Aplica estilos al texto en la celda B$totalRowSecond
        $sheet->getStyle('B' . $endRowSecond)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle('B' . $endRowSecond)->getFont()->setSize(11);

        // Escribe el texto en la celda B$totalRowSecond
        $sheet->setCellValue('B' . $endRowSecond, 'Total');

        // Calcular la suma de los valores en las celdas E$startRowSecond hasta E$endRowSecond
        $sumaDebeSecond = '=SUM(E' . $startRowSecond . ':E' . $endRowSecond . ')';

        // Asignar la fórmula de suma a la celda E$totalRowSecond
        $sheet->setCellValue('E' . $endRowSecond, $sumaDebeSecond . ' €');

        // Aplicar el formato numérico a la celda E$totalRowSecond
        $sheet->getStyle('E' . $endRowSecond)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        // Calcular la suma de los valores en las celdas F$startRowSecond hasta F$endRowSecond
        $sumaHaberSecond = '=SUM(F' . $startRowSecond . ':F' . $endRowSecond . ')';

        // Asignar la fórmula de suma a la celda F$totalRowSecond
        $sheet->setCellValue('F' . $endRowSecond, $sumaHaberSecond . ' €');

        // Aplicar el formato numérico a la celda F$totalRowSecond
        $sheet->getStyle('F' . $endRowSecond)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');

        // Calcular la resta de los resultados en las celdas E$totalRowSecond y F$totalRowSecond
        $restaSaldoSecond = '=E' . $endRowSecond . '-F' . $endRowSecond;

        // Asignar la fórmula de resta a la celda G$totalRowSecond
        $sheet->setCellValue('G' . $endRowSecond, $restaSaldoSecond . ' €');

        // Aplicar el formato numérico a la celda G$totalRowSecond
        $sheet->getStyle('G' . $endRowSecond)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_NUMBER_00 . '€');
    }
 */
    public function styles(Worksheet $sheet)
    {
        // Llamar a la función clearRow1
        $this->clearRow1($sheet);
        // Llamar a la función clearRow2
        $this->clearRow2($sheet);
        // Llamar a la función clearRow3
        $this->clearRow3($sheet);
        // Llamar a la función clearRow4
        $this->clearRow4($sheet);

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
        $endRow = 50; // Última fila donde se insertarán datos

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
        
        // Llamar a insertTotal para la primera sección de datos
        //$this->insertTotal($sheet, $startRow, $endRow, $totalDebe, $totalHaber);

        // Llamar a insertTotalSecond para la segunda sección de datos
        //$this->insertTotalSecond($sheet, $startRow, $endRow, $totalDebe, $totalHaber);
    }
}