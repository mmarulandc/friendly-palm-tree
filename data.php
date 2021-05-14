<?php

return [
    "Cliente" => [
        "Empresa" => "Ladrillos y Paredes",
        "Encargado" => "Juanito Gomez",
        "Logo" => "https://somedomain.com/pew.png",
    ],
    "Matriz" => [
        "Columnas" => [
            "Column1",
            "Column2",
            "Column3",
        ],
        "Cumplimiento" => [
            [1 => "Cumple"],
            [2 => "Cumple parcialmente"],
            [4 => "No cumple"],
            [5 => "Informativo"],
            [6 => "Informativo"]
        ],
        "Values" => [
            1 => [
                "Column1" => "Some text a1",
                "Column2" => "Some text2 a1",
                "Column3" => "Some text3 a3",
            ],
            2 => [
                "Column1" => "Some text b1",
                "Column2" => "Some text2 b2",
                "Column3" => "Some text3 b3",
            ],
            3 => [
                "Column1" => "Some text c1",
                "Column2" => "Some text2 c2",
                "Column3" => "Some text3 c3",
            ],
            4 => [
                "Column1" => "Some text d1",
                "Column2" => "Some text2 d2",
                "Column3" => "Some text3 d3",
            ],
            5 => [
                "Column1" => "Some text e1",
                "Column2" => "Some text2 e2",
                "Column3" => "Some text3 e3",
            ]
        ]
    ]
];

// function getCumplimientos($cumplimiento) {
//     return $cumplimiento[array_keys($cumplimiento)[0]];
// }
// function getCumplimientos($cumplimiento) {
//     return $cumplimiento[array_keys($cumplimiento)[0]];
//   }
$array = array_map('getCumplimientos', $data['Matriz']['Values']);
var_dump($array);
// $cumplimientos = $data['Matriz']["Cumplimiento"];
// foreach ($cumplimientos as &$cumplimiento) {
//     $key = array_keys($cumplimiento);
//     var_dump($cumplimiento[$key[0]]);
// }
// var_dump($data['Matriz']['Cumplimiento'][0][1]);

