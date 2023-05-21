<?php
include('koneksi.inc.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet    = new Spreadsheet();
$sheet          = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1','namaayah');
$sheet->setCellValue('B1','tlayah');
$sheet->setCellValue('C1','pendayah');
$sheet->setCellValue('D1','pekerayah');
$sheet->setCellValue('E1','salaryayah');
$sheet->setCellValue('F1','berkebayah');

$query = mysqli_query($koneksi,"select * from dataayah");
$i  = 2;

while($row = mysqli_fetch_array($query)) 
{
    $sheet->setCellValue('A'.$i,$row['namaayah']);
    $sheet->setCellValue('B'.$i,$row['tlayah']);
    $sheet->setCellValue('C'.$i,$row['pendayah']);
    $sheet->setCellValue('D'.$i,$row['pekerayah']);
    $sheet->setCellValue('E'.$i,$row['salaryayah']);
    $sheet->setCellValue('F'.$i,$row['berkebayah']);
    $i++;
}

$styleArray = [
    'borders'=>[
        'allBorders'=>[
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$i = $i - 1;
$sheet->getStyle('A1:F'.$i)->applyFromArray($styleArray);

$writer         = new Xlsx($spreadsheet);
$writer->save('Report Data Ayah.Xlsx');
if ($writer) { 
    header("location: prosesdatayah.php"); 
} 
?>