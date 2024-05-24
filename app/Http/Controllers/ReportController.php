<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use Symfony\Component\HttpFoundation\StreamedResponse;


class ReportController extends Controller {
    
    public function index() {
        $spreadsheet = new Spreadsheet();
        $objPHPExcel = $spreadsheet->getActiveSheet();

        $objPHPExcel->getStyle('A1')->applyFromArray(
            array ('font' => array('size' => 11, 'bold' => true, 'color' => array('rgb' => '000000')))
        );

        $styleArray2 = array(                   
            'font'  => array(
                'bold'  => true,
                'size'  => 14
            )
        );

        $styleArray3 = array(                   
            'font'  => array(
                'bold'  => true,
                'size'  => 11
            )
        );
        
        // Apply styles to header row
        $objPHPExcel->getRowDimension(1)->setRowHeight(20);
        $objPHPExcel->getStyle('A1:E1')->applyFromArray($styleArray2);

        // Set header values
        $objPHPExcel->SetCellValue('A1', "Name");
        $objPHPExcel->SetCellValue('B1', "Categories");
        $objPHPExcel->SetCellValue('C1', "Task Description");
        $objPHPExcel->SetCellValue('D1', "Date of Entry");
        $objPHPExcel->SetCellValue('E1', "Comments");

        // Fetch user data
        $user = $this->data();

        // Initialize starting row ID
        $rowId = 2;

        // Loop through each user and set cell values using the private method
        foreach ($user as $row) {
            $this->setCellValue($objPHPExcel, $row, $rowId);
            $rowId++;
        }

        // Set column widths
        $objPHPExcel->getColumnDimension('A')->setWidth(30);
        $objPHPExcel->getColumnDimension('B')->setWidth(30);
        $objPHPExcel->getColumnDimension('C')->setWidth(30);
        $objPHPExcel->getColumnDimension('D')->setWidth(30);
        $objPHPExcel->getColumnDimension('E')->setWidth(30);
        
        $writer = new Xls($spreadsheet);
        $response = new StreamedResponse(function() use ($writer) {
            $writer->save('php://output');
        });

        $response->headers->set('Content-Type', 'application/vnd.ms-excel');
        $response->headers->set('Content-Disposition', 'attachment;filename="data.xls"');
        $response->headers->set('Cache-Control', 'max-age=0');

        return $response;
    }


    private function setCellValue($objPHPExcel, $row, $rowId) {
        /*check if employee has daily performance data*/
        if(!isset($row['daily_performances']) || empty($row['daily_performances'])) return;

        /*employee name*/
        $objPHPExcel->SetCellValue('A'.$rowId, $row['name']);
        
        /*employee categories*/
        $categories = !empty($row['categories']) ? implode(", ", $row['categories']) : "";
        $objPHPExcel->SetCellValue('B'.$rowId, $categories);

        /*employee daily performance data*/
        $tasks = [];
        $datetime = "";
        $comment = "";
        
        if(isset($row['daily_performances']) && !empty($row['daily_performances'])){
            foreach($row['daily_performances'] as $performance){
                $tasks[]    = $performance['task']['name'] ?? ""; //task name
                $datetime   = $performance['datetime'] ?? "";    //datetime
                $comment    = $performance['comment'] ?? "";    //comment
            }
        }

        $objPHPExcel->SetCellValue('C'.$rowId, implode(", ", $tasks)); 
        $objPHPExcel->SetCellValue('D'.$rowId, $datetime);
        $objPHPExcel->SetCellValue('E'.$rowId, $comment);
    }

    private function data(){
        $data = [
            [
                "id" => 2,
                "category_ids" => "[\"3\"]",
                "branch_name" => "IT Departement",
                "name" => "Vijay",
                "email" => "vijay.g.php@gmail.com",
                "email_verified_at" => null,
                "role" => "employees",
                "status" => "active",
                "deleted_at" => null,
                "created_at" => "2024-05-21T11:17:34.000000Z",
                "updated_at" => "2024-05-22T12:21:57.000000Z",
                "categories" => [
                    "Information Technology"
                ],
                "full_name" => "Vijay (vijay.g.php@gmail.com)",
                "daily_performances" => [
                    [
                        "id" => 35,
                        "user_id" => 2,
                        "task_id" => 8,
                        "category_id" => "[\"3\"]",
                        "datetime" => "2024-07-26 11:59:30",
                        "comment" => "excellent",
                        "created_at" => "2024-05-21T12:00:00.000000Z",
                        "updated_at" => "2024-05-21T12:00:00.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 8,
                            "name" => "Quality Control",
                            "category_ids" => "[\"1\",\"3\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:15:23.000000Z",
                            "updated_at" => "2024-05-21T11:15:23.000000Z"
                        ]
                    ],
                    [
                        "id" => 36,
                        "user_id" => 2,
                        "task_id" => 9,
                        "category_id" => "[\"3\"]",
                        "datetime" => "2024-07-31 11:05:30",
                        "comment" => "up to par",
                        "created_at" => "2024-05-21T12:00:00.000000Z",
                        "updated_at" => "2024-05-21T12:00:00.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 9,
                            "name" => "Server Management",
                            "category_ids" => "[\"3\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:16:03.000000Z",
                            "updated_at" => "2024-05-21T11:16:03.000000Z"
                        ]
                    ]
                ]
            ],
            [
                "id" => 3,
                "category_ids" => "[\"1\"]",
                "branch_name" => "Electrical",
                "name" => "John",
                "email" => "John.g.php@gmail.com",
                "email_verified_at" => null,
                "role" => "employees",
                "status" => "active",
                "deleted_at" => null,
                "created_at" => "2024-05-21T11:18:26.000000Z",
                "updated_at" => "2024-05-21T11:18:26.000000Z",
                "categories" => [
                    "Electrical"
                ],
                "full_name" => "John (John.g.php@gmail.com)",
                "daily_performances" => [
                    [
                        "id" => 37,
                        "user_id" => 3,
                        "task_id" => 3,
                        "category_id" => "[\"1\"]",
                        "datetime" => "2024-06-27 12:04:05",
                        "comment" => "Took long on this task",
                        "created_at" => "2024-05-21T12:00:40.000000Z",
                        "updated_at" => "2024-05-21T12:00:40.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 3,
                            "name" => "Maintaining",
                            "category_ids" => "[\"1\",\"2\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:12:56.000000Z",
                            "updated_at" => "2024-05-21T11:12:56.000000Z"
                        ]
                    ],
                    [
                        "id" => 38,
                        "user_id" => 3,
                        "task_id" => 8,
                        "category_id" => "[\"1\"]",
                        "datetime" => "2024-06-01 18:00:05",
                        "comment" => "resource wasted",
                        "created_at" => "2024-05-21T12:00:40.000000Z",
                        "updated_at" => "2024-05-21T12:00:40.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 8,
                            "name" => "Quality Control",
                            "category_ids" => "[\"1\",\"3\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:15:23.000000Z",
                            "updated_at" => "2024-05-21T11:15:23.000000Z"
                        ]
                    ]
                ]
            ],
            [
                "id" => 4,
                "category_ids" => "[\"2\"]",
                "branch_name" => "Mechanical",
                "name" => "Tom",
                "email" => "tom.g.php@gmail.com",
                "email_verified_at" => null,
                "role" => "employees",
                "status" => "active",
                "deleted_at" => null,
                "created_at" => "2024-05-21T11:19:00.000000Z",
                "updated_at" => "2024-05-21T11:19:00.000000Z",
                "categories" => [
                    "Mechanical"
                ],
                "full_name" => "Tom (tom.g.php@gmail.com)",
                "daily_performances" => [
                    [
                        "id" => 33,
                        "user_id" => 4,
                        "task_id" => 2,
                        "category_id" => "[\"2\"]",
                        "datetime" => "2024-07-25 15:58:40",
                        "comment" => "Improve",
                        "created_at" => "2024-05-21T11:59:01.000000Z",
                        "updated_at" => "2024-05-21T11:59:01.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 2,
                            "name" => "Troubleshooting",
                            "category_ids" => "[\"1\",\"2\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:12:43.000000Z",
                            "updated_at" => "2024-05-21T11:12:43.000000Z"
                        ]
                    ],
                    [
                        "id" => 34,
                        "user_id" => 4,
                        "task_id" => 3,
                        "category_id" => "[\"2\"]",
                        "datetime" => "2024-08-29 16:58:40",
                        "comment" => "This is nice",
                        "created_at" => "2024-05-21T11:59:01.000000Z",
                        "updated_at" => "2024-05-21T11:59:01.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 3,
                            "name" => "Maintaining",
                            "category_ids" => "[\"1\",\"2\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:12:56.000000Z",
                            "updated_at" => "2024-05-21T11:12:56.000000Z"
                        ]
                    ]
                ]
            ],
            [
                "id" => 5,
                "category_ids" => "[\"3\",\"2\"]",
                "branch_name" => "Multiple",
                "name" => "Raj",
                "email" => "raj.g.php@gmail.com",
                "email_verified_at" => null,
                "role" => "employees",
                "status" => "active",
                "deleted_at" => null,
                "created_at" => "2024-05-21T11:19:44.000000Z",
                "updated_at" => "2024-05-21T11:19:44.000000Z",
                "categories" => [
                    "Mechanical",
                    "Information Technology"
                ],
                "full_name" => "Raj (raj.g.php@gmail.com)",
                "daily_performances" => []
            ],
            [
                "id" => 6,
                "category_ids" => "[\"1\",\"3\",\"2\"]",
                "branch_name" => "Mutiple",
                "name" => "Paul",
                "email" => "paul.g.php@gmail.com",
                "email_verified_at" => null,
                "role" => "employees",
                "status" => "active",
                "deleted_at" => null,
                "created_at" => "2024-05-21T11:21:23.000000Z",
                "updated_at" => "2024-05-21T11:21:23.000000Z",
                "categories" => [
                    "Electrical",
                    "Mechanical",
                    "Information Technology"
                ],
                "full_name" => "Paul (paul.g.php@gmail.com)",
                "daily_performances" => [
                    [
                        "id" => 2,
                        "user_id" => 6,
                        "task_id" => 2,
                        "category_id" => "[\"1\",\"3\",\"2\"]",
                        "datetime" => "2024-05-21 11:21:28",
                        "comment" => "You have done well, but improve more",
                        "created_at" => "2024-05-21T11:22:03.000000Z",
                        "updated_at" => "2024-05-21T12:18:59.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 2,
                            "name" => "Troubleshooting",
                            "category_ids" => "[\"1\",\"2\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:12:43.000000Z",
                            "updated_at" => "2024-05-21T11:12:43.000000Z"
                        ]
                    ],
                    [
                        "id" => 3,
                        "user_id" => 6,
                        "task_id" => 3,
                        "category_id" => "[\"1\",\"3\",\"2\"]",
                        "datetime" => "2024-08-15 14:21:28",
                        "comment" => "Improve your efficient",
                        "created_at" => "2024-05-21T11:22:03.000000Z",
                        "updated_at" => "2024-05-21T12:02:37.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 3,
                            "name" => "Maintaining",
                            "category_ids" => "[\"1\",\"2\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:12:56.000000Z",
                            "updated_at" => "2024-05-21T11:12:56.000000Z"
                        ]
                    ],
                    [
                        "id" => 4,
                        "user_id" => 6,
                        "task_id" => 4,
                        "category_id" => "[\"1\",\"3\",\"2\"]",
                        "datetime" => "2024-08-17 11:25:31",
                        "comment" => "Task comment updated",
                        "created_at" => "2024-05-21T11:22:03.000000Z",
                        "updated_at" => "2024-05-21T12:02:18.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 4,
                            "name" => "installation",
                            "category_ids" => "[\"1\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:13:24.000000Z",
                            "updated_at" => "2024-05-21T11:13:24.000000Z"
                        ]
                    ],
                    [
                        "id" => 5,
                        "user_id" => 6,
                        "task_id" => 5,
                        "category_id" => "[\"1\",\"3\",\"2\"]",
                        "datetime" => "2024-10-30 11:25:28",
                        "comment" => "Good update",
                        "created_at" => "2024-05-21T11:22:03.000000Z",
                        "updated_at" => "2024-05-21T13:20:51.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 5,
                            "name" => "CAD Modeling",
                            "category_ids" => "[\"2\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:14:35.000000Z",
                            "updated_at" => "2024-05-21T11:14:35.000000Z"
                        ]
                    ],
                    [
                        "id" => 31,
                        "user_id" => 6,
                        "task_id" => 1,
                        "category_id" => "[\"1\",\"3\",\"2\"]",
                        "datetime" => "2024-08-22 15:03:01",
                        "comment" => "Task was done good",
                        "created_at" => "2024-05-21T11:58:35.000000Z",
                        "updated_at" => "2024-05-21T12:01:55.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 1,
                            "name" => "Designing",
                            "category_ids" => "[\"1\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:12:26.000000Z",
                            "updated_at" => "2024-05-21T11:12:26.000000Z"
                        ]
                    ],
                    [
                        "id" => 32,
                        "user_id" => 6,
                        "task_id" => 2,
                        "category_id" => "[\"1\",\"3\",\"2\"]",
                        "datetime" => "2024-09-05 11:08:01",
                        "comment" => "Very nice",
                        "created_at" => "2024-05-21T11:58:35.000000Z",
                        "updated_at" => "2024-05-21T12:01:46.000000Z",
                        "deleted_at" => null,
                        "task" => [
                            "id" => 2,
                            "name" => "Troubleshooting",
                            "category_ids" => "[\"1\",\"2\"]",
                            "status" => "active",
                            "deleted_at" => null,
                            "created_at" => "2024-05-21T11:12:43.000000Z",
                            "updated_at" => "2024-05-21T11:12:43.000000Z"
                        ]
                    ]
                ]
            ]
        ];

        return $data;
    }    
}
