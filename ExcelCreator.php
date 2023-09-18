<?php

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require 'vendor/autoload.php';
require 'Decoder.php';

class ExcelCreator extends Decoder
{
    private string $outputPath;

    public function __construct(string $outputPath)
    {
        $this->outputPath = $outputPath;
    }

    public function createExcelFromData(string $data): void
    {
        try {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();

            // Указываем заголовки столбцов
            $columnHeaders = [
                'Название',
                'Начало активности',
                'Команда 1',
                'Команда 2',
                'ID',
                'Матч SUPER CUP',
                'Очки команды 1',
                'Очки команды 2',
                'Подпись к матчу',
                'Результат матча',
                'Дата окончания матча',
            ];

            // Устанавливаем ширину столбцов
            $columnWidths = [
                'A' => 35,
                'B' => 20,
                'C' => 15,
                'D' => 15,
                'E' => 15,
                'F' => 25,
                'G' => 22,
                'H' => 22,
                'I' => 22,
                'J' => 22,
                'K' => 27,
            ];

            foreach ($columnHeaders as $key => $header) {
                $columnLetter = Coordinate::stringFromColumnIndex($key + 1);
                $sheet->setCellValue($columnLetter . '1', $header);
                $sheet->getColumnDimension($columnLetter)->setWidth($columnWidths[$columnLetter]);
            }

            // Применяем метод unserializeArray к данным из файла
            $dataArray = $this::unserializeArray($data);

            // Заполняем таблицу данными из массива
            $rowIndex = 2; // Начинаем с второй строки после заголовков
            foreach ($dataArray as $row) {
                if ($rowIndex === 2) {
                    foreach ($columnHeaders as $key => $header) {
                        $columnLetter = Coordinate::stringFromColumnIndex($key + 1);
                        $sheet->getStyle($columnLetter . '1')->getFont()->setBold(true)->setSize(12);
                    }
                }
                $sheet->setCellValue('A' . $rowIndex, $row['NAME']);
                $sheet->setCellValue('B' . $rowIndex, $row['ACTIVE_FROM']);
                $sheet->setCellValue('C' . $rowIndex, $row['PROPERTIES']['TEAM_1']['VALUE'] ?? '');
                $sheet->setCellValue('D' . $rowIndex, $row['PROPERTIES']['TEAM_2']['VALUE'] ?? '');
                $sheet->setCellValue('E' . $rowIndex, $row['ID']);
                $sheet->setCellValue('F' . $rowIndex, $row['PROPERTIES']['IS_DRAGON_CUP']['VALUE'] ?? '');
                $sheet->setCellValue('G' . $rowIndex, $row['PROPERTIES']['POINTS_TEAM_1']['VALUE'] ?? '');
                $sheet->setCellValue('H' . $rowIndex, $row['PROPERTIES']['POINTS_TEAM_2']['VALUE'] ?? '');
                $sheet->setCellValue('I' . $rowIndex, $row['PROPERTIES']['DOP_INFO_FOR_MATCH']['VALUE'] ?? '');
                $sheet->setCellValue('J' . $rowIndex, $row['PROPERTIES']['RESULT']['VALUE'] ?? '');
                $sheet->setCellValue('K' . $rowIndex, $row['PROPERTIES']['DATE_END']['VALUE'] ?? '');
                $rowIndex++;
            }

            // Получение даты, которая будет использоваться в имени файла
            $dt = date('Y-m-d_H-i-s');

            // Отправка файла в браузер
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment; filename=file-' . $dt . '.xlsx');

            $writer = new Xlsx($spreadsheet);
            $writer->save('php://output');
        } catch (Exception $e) {
            echo "Произошла ошибка: " . $e->getMessage();
        }
    }
}

$filename = 'keydata.txt';
$data = file_get_contents($filename);

if ($data === false) {
    echo "Не удалось прочитать файл.";
} else {
    try {
        $excelCreator = new ExcelCreator('./output/');
        $excelCreator->createExcelFromData($data);
    } catch (Exception $e) {
        echo "Произошла ошибка: " . $e->getMessage();

    }
}
