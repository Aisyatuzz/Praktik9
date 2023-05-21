<?php
include('koneksi.inc.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet    = new Spreadsheet();
$sheet          = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1','nopendaftaran');
$sheet->setCellValue('B1','jenispend');
$sheet->setCellValue('C1','tglmsksklh');
$sheet->setCellValue('D1','nis');
$sheet->setCellValue('E1','nopesujian');
$sheet->setCellValue('F1','appaud');
$sheet->setCellValue('G1','aptk');
$sheet->setCellValue('H1','noserskhun');
$sheet->setCellValue('I1','noserijazah');
$sheet->setCellValue('J1','hobi');
$sheet->setCellValue('K1','citacita');

$query = mysqli_query($koneksi,"select * from registrasi");
$i  = 2;
$no = 1;
while($row = mysqli_fetch_array($query)) 
{
    $sheet->setCellValue('A'.$i,$no++);
    $sheet->setCellValue('B'.$i,$row['jenispend']);
    $sheet->setCellValue('C'.$i,$row['tglmsksklh']);
    $sheet->setCellValue('D'.$i,$row['nis']);
    $sheet->setCellValue('E'.$i,$row['nopesujian']);
    $sheet->setCellValue('F'.$i,$row['appaud']);
    $sheet->setCellValue('G'.$i,$row['aptk']);
    $sheet->setCellValue('H'.$i,$row['noserskhun']);
    $sheet->setCellValue('I'.$i,$row['noserijazah']);
    $sheet->setCellValue('J'.$i,$row['hobi']);
    $sheet->setCellValue('K'.$i,$row['citacita']);
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
$sheet->getStyle('A1:K'.$i)->applyFromArray($styleArray);

$writer         = new Xlsx($spreadsheet);
$writer->save('Report Data Registrasi.Xlsx');
if ($writer) { 
    header("location: prosesregistrasi.php"); 
} 
?>