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
        // Borra los datos desde la celda E7 hasta la celda G7
        for ($column = 'E'; $column <= 'G'; $column++) {
            $sheet->setCellValue($column . '7', '');
        }
    }

    public function addTotal(Worksheet $sheet, $startRow, $endRow, $startColumn, $endColumn)
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
        $totalRow = $lastRowWithData + 2;
        // Verificar si la fila total está dentro del rango, ajustar si es necesario
        if ($totalRow > $endRow) {
            $totalRow = $endRow + 2;
        }
        $totalColumn = 'B'; // Columna B
        $sheet->getStyle($totalColumn . $totalRow)->getFont()->setBold(true);
        $sheet->getStyle($totalColumn . $totalRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue($totalColumn . $totalRow, "Total:");

        // Definir variables dentro del bucle para reiniciarlas en cada iteración
        $sumaColumnaDebe = 0;
        $sumaColumnaHaber = 0;

        // Calcular las sumas de la columna "Debe" y "Haber"
        foreach ($this->collection() as $japanta) {
            $sumaColumnaDebe += $japanta->debe;
            $sumaColumnaHaber += $japanta->haber;
        }

        // Calcular el saldo
        $saldo = $sumaColumnaDebe - $sumaColumnaHaber;

        // Insertar dos filas debajo de la fila "Total"
        $sheet->insertNewRowBefore($totalRow + 1, 2);

        // Escribir los valores en las nuevas filas
        $sheet->setCellValue("E{$totalRow}", $sumaColumnaDebe . ' €');
        $sheet->setCellValue("F{$totalRow}", $sumaColumnaHaber . ' €');
        $sheet->setCellValue("G{$totalRow}", $saldo . ' €');

        // Aplicar alineación a la izquierda en las celdas
        $sheet->getStyle("E{$totalRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle("F{$totalRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle("G{$totalRow}")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    }

    public function addIVA10(Worksheet $sheet, $startRow, $endRow, $startColumn, $endColumn)
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

        // Agregar "4720000010, Hp, IVA Soportado 10%" en negrita a la izquierda de una fila debajo del último dato
        $totalRow = $lastRowWithData + 2;
        // Verificar si la fila total está dentro del rango, ajustar si es necesario
        if ($totalRow > $endRow) {
            $totalRow = $endRow + 2;
        }
        $totalColumn = 'A'; // Columna A

        // Combinar celdas A y B
        $sheet->mergeCells($totalColumn . $totalRow . ':' . 'B' . $totalRow);

        $sheet->getStyle($totalColumn . $totalRow)->getFont()->setBold(true);
        $sheet->getStyle($totalColumn . $totalRow)->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        $sheet->setCellValue($totalColumn . $totalRow, "4720000010, Hp, IVA Soportado 10%");

        // Agregar "Fecha" en negrita a la izquierda de una fila debajo del último dato
        $totalRow = $lastRowWithData + 3;
        // Verificar si la fila total está dentro del rango, ajustar si es necesario
        if ($totalRow > $endRow) {
            $totalRow = $endRow + 3;
        }

        // Iterar sobre las columnas desde la A hasta la G
        for ($column = 'A'; $column <= 'G'; $column++) {
            // Obtener el estilo de la celda actual
            $style = $sheet->getStyle($column . $totalRow);

            // Aplicar el estilo de fuente negrita y color blanco
            $style->getFont()->setBold(true)->setColor(new \PhpOffice\PhpSpreadsheet\Style\Color(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE));

            // Aplicar el color de fondo negro
            $style->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
            $style->getFill()->getStartColor()->setRGB('000000');

            // Aplicar la alineación horizontal a la izquierda
            $style->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
        }

        // Obtener la celda A en la fila específica
        $cellA = 'A' . $totalRow;
        // Asignar el texto "Fecha" a la celda A
        $sheet->setCellValue($cellA, "Fecha");

        // Obtener la celda B en la fila específica
        $cellB = 'B' . $totalRow;
        // Asignar el texto "Concepto" a la celda B
        $sheet->setCellValue($cellB, "Concepto");

        // Obtener la celda C en la fila específica
        $cellC = 'C' . $totalRow;
        // Asignar el texto "Documento" a la celda C
        $sheet->setCellValue($cellC, "Documento");

        // Obtener la celda D en la fila específica
        $cellD = 'D' . $totalRow;
        // Asignar el texto "Tags" a la celda D
        $sheet->setCellValue($cellD, "Tags");

        // Obtener la celda E en la fila específica
        $cellE = 'E' . $totalRow;
        // Asignar el texto "Debe" a la celda E
        $sheet->setCellValue($cellE, "Debe");

        // Obtener la celda F en la fila específica
        $cellF = 'F' . $totalRow;
        // Asignar el texto "Haber" a la celda F
        $sheet->setCellValue($cellF, "Haber");

        // Obtener la celda G en la fila específica
        $cellG = 'G' . $totalRow;
        // Asignar el texto "Saldo" a la celda G
        $sheet->setCellValue($cellG, "Saldo");
    }

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

        // Definir variables fuera del bucle para no reiniciarlas en cada iteración
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
        
        // Llamar a addTotal para la primera sección de datos
        $this->addTotal($sheet, $startRow, $endRow, 'B', 'G');

        // Llamar a addIVA10 para la primera sección de datos
        $this->addIVA10($sheet, $startRow, $endRow, 'A', 'G');
    }
}