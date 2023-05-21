<?php
include('koneksi.inc.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet    = new Spreadsheet();
$sheet          = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1','namaibu');
$sheet->setCellValue('B1','tlibu');
$sheet->setCellValue('C1','pendibu');
$sheet->setCellValue('D1','pekeribu');
$sheet->setCellValue('E1','salaryibu');
$sheet->setCellValue('F1','berkebibu');

$query = mysqli_query($koneksi,"select * from dataibu");
$i  = 2;

while($row = mysqli_fetch_array($query)) 
{
    $sheet->setCellValue('A'.$i,$row['namaibu']);
    $sheet->setCellValue('B'.$i,$row['tlibu']);
    $sheet->setCellValue('C'.$i,$row['pendibu']);
    $sheet->setCellValue('D'.$i,$row['pekeribu']);
    $sheet->setCellValue('E'.$i,$row['salaryibu']);
    $sheet->setCellValue('F'.$i,$row['berkebibu']);
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
$writer->save('Report Data Ibu.Xlsx');
if ($writer) { 
    header("location: prosesdataibu.php"); 
} 
?>