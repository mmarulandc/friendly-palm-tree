<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

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
  $drawing = new Drawing();
  $drawing->setName('Company Logo');
  $drawing->setDescription('Company Logo');
  $drawing->setPath($logo);
  $drawing->setWidthAndHeight(90, 90);
  $drawing->setOffsetX(70);
  $drawing->setOffsetY(5);
  $drawing->setCoordinates('B2');
  $drawing->setWorksheet($spreadsheet->getActiveSheet());
  $spreadsheet->getActiveSheet()->setCellValue('C3', $companyName);
  $spreadsheet->getActiveSheet()->setCellValue('D3', $inCharge);


}

function getCumplimientos($cumplimiento) {
  return $cumplimiento[array_keys($cumplimiento)[0]];
}

function getCumplimientoKey($cumplimiento) {
  return array_keys($cumplimiento)[0];
}


function setCumplimientoData($cumplimientos, $spreadsheet) {
  $styles = [
    'fullfill'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'c6efce'
        ],
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '006100'
        ]
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'indent' => 1
      ]
    ],
    'partiallyFullfill'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'ffeb9c'
        ]
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '9c5600'
        ]
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'indent' => 1
      ]
    ],
    'notFullfill'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'ffc7ce'
        ]
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '9c0005'
        ]
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'indent' => 1
      ]
    ],
    'notApply'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'd9e1f2'
        ]
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '4472C4'
        ]
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'indent' => 1
      ]
    ],
    'pending'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'd9d9d9'
        ],

      ],
      'font' => [
        'name' => 'Arial Nova',
        
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'indent' => 1
      ]
    ],
    'informative'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'A5E6F1'
        ],
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '188698'
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ]
    ],
    'future'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'DFBBFD'
        ],
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '6805B9'
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ]
    ],
    'futureProject'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'EBD2EE'
        ],
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '9D40AA'
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ]
    ],
    'futureManagement'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'E2BFE7'
        ],
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => 'A844B6'
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ]
    ],
    'futureContingency'=> [
      'fill' => [
        'fillType' => Fill::FILL_SOLID,
        'color' => [
          'argb' => 'D6C0D3'
        ],
      ],
      'font' => [
        'name' => 'Arial Nova',
        'color' => [
          'argb' => '62405E'
        ]
      ],
      'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
      ],
      'borders' => [
        'bottom' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
          'color' => ['argb' => '000000'],
        ],
        'right' => [
          'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
          'color' => ['argb' => '000000'],
        ]
      ]
    ],
  ];
  $keysValues = array_chunk(array_map('getCumplimientoKey', $cumplimientos), 1);
  $spreadsheet->getActiveSheet()->fromArray($keysValues, NULL, 'A5');

  for($row = 5; $row !=  (5+count($keysValues)); $row++) {
    $spreadsheet->getActiveSheet()->duplicateStyle($spreadsheet->getActiveSheet()->getStyle('A5'), 'A'.$row);
    $spreadsheet->getActiveSheet()->getRowDimension($row)->setRowHeight(26);
  }

  $array = array_chunk(array_map('getCumplimientos', $cumplimientos), 1);
  $spreadsheet->getActiveSheet()->fromArray($array, NULL, 'B5');
  $maxRow = 5 + count($cumplimientos);
  // $spreadsheet->getActiveSheet()->fromArray($keysValues, NULL, 'A5');


  $conditionalStyles = $spreadsheet->getActiveSheet()
    ->getStyle('B5:B'.$maxRow)
    ->getConditionalStyles();

  $fullfillCondition = new Conditional();
  $partiallyFullfillCondition = new Conditional();
  $notFullfill = new Conditional();
  $pending = new Conditional();
  $notApply = new Conditional();
  $informative = new Conditional(); 
  $future = new Conditional(); 
  $futureProject = new Conditional(); 
  $futureManagement = new Conditional(); 
  $futureContingency = new Conditional(); 

  $fullfillCondition->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
    ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
    ->setText('Cumple');

  $fullfillCondition->getStyle()->applyFromArray($styles['fullfill']);
  
  $partiallyFullfillCondition->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
    ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
    ->setText('Cumple parcialmente');
  
  $partiallyFullfillCondition->getStyle()->applyFromArray($styles['partiallyFullfill']);

  $notFullfill->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
    ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
    ->setText('No cumple');
  $notFullfill->getStyle()->applyFromArray($styles['notFullfill']);

  $pending->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
  ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
  ->setText('Pendiente');
  $pending->getStyle()->applyFromArray($styles['pending']);
  
  
  $notApply->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
  ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
  ->setText('No aplica');
  $notApply->getStyle()->applyFromArray($styles['notApply']);

  $informative->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
  ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
  ->setText('Informativa');
  $informative->getStyle()->applyFromArray($styles['informative']);

  $futureProject->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
  ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
  ->setText('Futuro-proyecto');
  $futureProject->getStyle()->applyFromArray($styles['futureProject']);

  $futureManagement->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
  ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
  ->setText('Futuro-Gestión a corto plazo');
  $futureManagement->getStyle()->applyFromArray($styles['futureManagement']);

  $futureContingency->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
  ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
  ->setText('Futuro-contingencia');
  $futureContingency->getStyle()->applyFromArray($styles['futureContingency']);

  $future->setConditionType(Conditional::CONDITION_CONTAINSTEXT)
  ->setOperatorType(Conditional::OPERATOR_CONTAINSTEXT)
  ->setText('Futuro');
  $future->getStyle()->applyFromArray($styles['future']);

  array_push($conditionalStyles, $notFullfill);
  array_push($conditionalStyles, $partiallyFullfillCondition);
  array_push($conditionalStyles, $fullfillCondition);
  array_push($conditionalStyles, $pending);
  array_push($conditionalStyles, $notApply);
  array_push($conditionalStyles, $informative);
  array_push($conditionalStyles, $futureManagement);
  array_push($conditionalStyles, $futureContingency);
  array_push($conditionalStyles, $futureProject);
  array_push($conditionalStyles, $future);



  $spreadsheet->getActiveSheet()
    ->getStyle('B5:B'.$maxRow)
    ->setConditionalStyles($conditionalStyles);
}


function applyBorder($spreadsheet, $borderRow, $borderColumn) {
  $borderStyle = [
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
            'bottom' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
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
  }

}


function applyColorOffset($spreadsheet, $data) {
  $colorOffsetRigthStart = numberToExcelLetter(3 + count($data['Matriz']['Columnas'])).'4';
  $colorOffsetRigthEnd = numberToExcelLetter(4 + count($data['Matriz']['Columnas'])).(4 + count($data['Matriz']['Cumplimiento']));
  $colorOffsetBottomStart = 'B'.(5 + count($data['Matriz']['Cumplimiento']));
  $colorOffsetBottomEnd = numberToExcelLetter(4 + count($data['Matriz']['Columnas'])).(6 + count($data['Matriz']['Cumplimiento']));

  $offsetStyle = [
    'fill' => [
      'fillType' => Fill::FILL_SOLID,
      'color' => [
        'argb' => 'D8D8D8'
      ]
    ],
  ];
  $spreadsheet->getActiveSheet()->getStyle($colorOffsetRigthStart.':'.$colorOffsetRigthEnd)->applyFromArray($offsetStyle);
  $spreadsheet->getActiveSheet()->getStyle($colorOffsetBottomStart.':'.$colorOffsetBottomEnd)->applyFromArray($offsetStyle);
}


function generateXls($data) {
  $spreadsheet = readTemplate();
  setInitialData($data, $spreadsheet);
  setHeaders($data['Matriz']['Columnas'], $spreadsheet);
  setCumplimientoData($data['Matriz']['Cumplimiento'], $spreadsheet);
  $borderColumn = numberToExcelLetter(2 + count($data['Matriz']['Columnas']));
  $borderRow = 4 + count($data['Matriz']['Cumplimiento']);
  applyBorder($spreadsheet, $borderRow, $borderColumn );
  applyTableFormat($spreadsheet, $data['Matriz']['Values'], $data['Matriz']['Columnas']);
  applyColorOffset($spreadsheet, $data);


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

generateXls($data);

?>