<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use Symfony\Component\HttpFoundation\StreamedResponse;

class ExportController extends Controller{
    //composer require phpoffice/phpspreadsheet

    public function export_data(){
        $spreadsheet = new Spreadsheet();
        $objPHPExcel = $spreadsheet->getActiveSheet();
        
        $objPHPExcel->getStyle('A1')->applyFromArray
        (
            array ('font' => array('size' => 11,'bold' => true,'color' => array('rgb' => '000000')))
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
        
        // font bold & line height big
        $variable1 = array(1);
        foreach ($variable1 as $key1 => $value1) {
            $objPHPExcel->getRowDimension($value1)->setRowHeight(20);
            $objPHPExcel->getStyle('A'.$value1.':F'.$value1)->applyFromArray($styleArray2);    
        }

        $objPHPExcel->SetCellValue('A1', "ID");
        $objPHPExcel->SetCellValue('B1', "Name");
        $objPHPExcel->SetCellValue('C1', "Email");
        $objPHPExcel->SetCellValue('D1', "Order ID");
        $objPHPExcel->SetCellValue('E1', "Order Date");
        $objPHPExcel->SetCellValue('F1', "Account Details");

        $rowId = 2;
        $customer = $this->get_data();
        foreach ($customer as $row) {
            $objPHPExcel->SetCellValue('A'.$rowId, $row['customer_id']);
            $objPHPExcel->SetCellValue('B'.$rowId, $row['name']);
            $objPHPExcel->SetCellValue('C'.$rowId, $row['email']);
            $objPHPExcel->SetCellValue('D'.$rowId, $row['order_id']);
            $objPHPExcel->SetCellValue('E'.$rowId, $row['order_date']);
            $objPHPExcel->SetCellValue('F'.$rowId, isset($row['account_details']) ? implode(', ', $row['account_details']) : "");
            $rowId++;
        }

        $objPHPExcel->getColumnDimension('A')->setWidth(30);
        $objPHPExcel->getColumnDimension('B')->setWidth(30);
        $objPHPExcel->getColumnDimension('C')->setWidth(30);
        $objPHPExcel->getColumnDimension('D')->setWidth(30);
        $objPHPExcel->getColumnDimension('E')->setWidth(30);
        $objPHPExcel->getColumnDimension('F')->setWidth(30);
        
        $writer     = new Xls($spreadsheet);
        $response   = new StreamedResponse(function() use ($writer) {
            $writer->save('php://output');
        });

        $response->headers->set('Content-Type', 'application/vnd.ms-excel');
        $response->headers->set('Content-Disposition', 'attachment;filename="data.xls"');
        $response->headers->set('Cache-Control', 'max-age=0');

        return $response;
    }

   private function get_data() {
        return [
            [
                "customer_id" => 1,
                "name" => "Alice",
                "email" => "alice@example.com",
                "order_id" => 101,
                "order_date" => "2023-01-15",
                "amount" => 250.00,
                "account_details" => [
                    "account_id" => 1055,
                    "account_status" => "Active"
                ],
                "shipping_address" => [
                    "street" => "123 Main St",
                    "city" => "Springfield",
                    "zipcode" => "12345"
                ]
            ],
            [
                "customer_id" => 2,
                "name" => "Bob",
                "email" => "bob@example.com",
                "order_id" => 102,
                "order_date" => "2023-01-17",
                "amount" => 150.00,
                "account_details" => [
                    "account_id" => 1056,
                    "account_status" => "Active"
                ],
                "preferences" => [
                    "newsletter" => true,
                    "sms_alerts" => false
                ]
            ],
            [
                "customer_id" => 3,
                "name" => "Charlie",
                "email" => "charlie@example.com",
                "order_id" => 104,
                "order_date" => "2023-02-01",
                "amount" => 450.00,
                "account_details" => [
                ],
                "referral" => [
                    "referred_by" => "Dave",
                    "referral_code" => "XYZ123"
                ]
            ]
        ];
    }
}
