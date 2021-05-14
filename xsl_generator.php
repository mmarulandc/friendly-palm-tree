<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\IOFactory;

//load template


// add content

//remove empty rows and columns

//conditional formatting
// ini_set('display_errors', 1);
// error_reporting(E_ALL);

function readTemplate() {
  $reader = IOFactory::createReader('Xlsx');
  $spreadsheet = $reader->load('./template/template_modified.xlsx');
  return $spreadsheet;
}

function numberToExcelLetter($number) {
  return chr(64 + $number);
}


function setHeaders($headersArray, $spreadsheet) {
  $excelColumnLetter = numberToExcelLetter(3 + count($headersArray));
  $spreadsheet->getActiveSheet()->fromArray($headersArray, NULL, 'C4');
  for($column = 'C'; $column != $excelColumnLetter; $column++) {
    $spreadsheet->getActiveSheet()->duplicateStyle($spreadsheet->getActiveSheet()->getStyle('C4'), $column.'4');
  }

}

function setInitialData($clientData, $spreadsheet) {
  $client = $clientData['Cliente'];
  $companyName = $client['Empresa'];
  $inCharge = $client['Encargado'];
  $logo = $client['Logo'];
  $spreadsheet->getActiveSheet()->setCellValue('B2', $logo);
  $spreadsheet->getActiveSheet()->setCellValue('C3', $companyName);
  $spreadsheet->getActiveSheet()->setCellValue('D3', $inCharge);

}

function getCumplimientos($cumplimiento) {
  return $cumplimiento[array_keys($cumplimiento)[0]];
}

// function getCumplimientos($cumplimiento) {
//   return $cumplimiento[array_keys($cumplimiento)[0]];
// }

function setCumplimientoData($cumplimientos, $spreadsheet) {
  $array = array_chunk(array_map('getCumplimientos', $cumplimientos), 1);
  $spreadsheet->getActiveSheet()->fromArray($array, NULL, 'B5');
}

function applyBorder($spreadsheet, $borderRow, $borderColumn) {
  $borderStyle = [
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
            'color' => ['argb' => '000000'],
        ],
    ],
  ];
  $spreadsheet->getActiveSheet()->getStyle('B4:'.$borderColumn.$borderRow)->applyFromArray($borderStyle);
}

function applyTableFormat ($spreadsheet, $rows, $columns) {
  $evenRow = [
    'fill' => [
      'fillType' => Fill::FILL_SOLID,
      'startColor' => [
        'argb' => 'ddebf7'
      ]
    ],
    'font' => [
      'name' => 'Arial Nova Light'
    ],
    'alignment' => [
      'horizontal' => Alignment::HORIZONTAL_LEFT,
      'vertical' => Alignment::VERTICAL_CENTER,
      'indent' => 1
    ],
  ];
  $oddRow = [
    'fill' => [
      'fillType' => Fill::FILL_SOLID,
      'startColor' => [
        'argb' => 'F5F5F5'
      ]
    ],
    'font' => [
      'name' => 'Arial Nova Light'
    ],
    'alignment' => [
      'horizontal' => Alignment::HORIZONTAL_LEFT,
      'vertical' => Alignment::VERTICAL_CENTER,
      'indent' => 1

    ],
  ];

  $spreadsheet->getActiveSheet()->fromArray(array_values($rows), NULL, 'C5');
  $highColumn = numberToExcelLetter(2 + count($columns)); // E
  var_dump(5 + count($rows));
  for($i=5; $i < (5 + count($rows)); $i++) {
    if ($i % 2 == 0) {
      $spreadsheet->getActiveSheet()->getStyle('C'.$i.':'.$highColumn . $i)->applyFromArray(
        $evenRow
      );
    } else {
      $spreadsheet->getActiveSheet()->getStyle('C'.$i.':'.$highColumn . $i)->applyFromArray(
        $oddRow
      );
    }
    var_dump('C'.$i.':'.$highColumn . $i);
  }

}

function generateXls($data) {
  // var_dump(is_array ( $style ) );
  $spreadsheet = readTemplate();
  setInitialData($data, $spreadsheet);
  setHeaders($data['Matriz']['Columnas'], $spreadsheet);
  setCumplimientoData($data['Matriz']['Cumplimiento'], $spreadsheet);
  $borderColumn = numberToExcelLetter(2 + count($data['Matriz']['Columnas']));
  $borderRow = 4 + count($data['Matriz']['Cumplimiento']);
  applyBorder($spreadsheet, $borderRow, $borderColumn );
  applyTableFormat($spreadsheet, $data['Matriz']['Values'], $data['Matriz']['Columnas']);
  // var_dump('B5:'.$endBorderColumn.$endBorderRow);
  //set the header first, so the result will be treated as an xlsx file.
  ob_end_clean();
  header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  header('Content-Disposition: attachment;filename="sheet.xlsx"' );
  header('Cache-Control: max-age=0');

  
  //create IOFactory object
  $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
  //save into php output
  $writer->save('php://output'); // download file

}





$data = require 'data.php';
// var_dump($styleArray);

generateXls($data);

?>