<?php
include('koneksi.inc.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet    = new Spreadsheet();
$sheet          = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1','namaleng');
$sheet->setCellValue('B1','jkelamin');
$sheet->setCellValue('C1','nisn');
$sheet->setCellValue('D1','nik');
$sheet->setCellValue('E1','tlahir');
$sheet->setCellValue('F1','tglahir');
$sheet->setCellValue('G1','agama');
$sheet->setCellValue('H1','berkebkhusus');
$sheet->setCellValue('I1','alamat');
$sheet->setCellValue('J1','rt');
$sheet->setCellValue('K1','rw');
$sheet->setCellValue('L1','namadusun');
$sheet->setCellValue('M1','namakel');
$sheet->setCellValue('N1','kec');
$sheet->setCellValue('O1','kodepos');
$sheet->setCellValue('P1','ttinggal');
$sheet->setCellValue('Q1','transport');
$sheet->setCellValue('R1','nohp');
$sheet->setCellValue('S1','notelp');
$sheet->setCellValue('T1','email');
$sheet->setCellValue('U1','kpspkh');
$sheet->setCellValue('V1','nokpspkh');
$sheet->setCellValue('W1','kwn');
$sheet->setCellValue('X1','namaneg');

$query = mysqli_query($koneksi,"select * from datapribadi");
$i  = 2;

while($row = mysqli_fetch_array($query)) 
{
    $sheet->setCellValue('A'.$i,$row['namaleng']);
    $sheet->setCellValue('B'.$i,$row['jkelamin']);
    $sheet->setCellValue('C'.$i,$row['nisn']);
    $sheet->setCellValue('D'.$i,$row['nik']);
    $sheet->setCellValue('E'.$i,$row['tlahir']);
    $sheet->setCellValue('F'.$i,$row['tglahir']);
    $sheet->setCellValue('G'.$i,$row['agama']);
    $sheet->setCellValue('H'.$i,$row['berkebkhusus']);
    $sheet->setCellValue('I'.$i,$row['alamat']);
    $sheet->setCellValue('J'.$i,$row['rt']);
    $sheet->setCellValue('K'.$i,$row['rw']);
    $sheet->setCellValue('L'.$i,$row['namadusun']);
    $sheet->setCellValue('M'.$i,$row['namakel']);
    $sheet->setCellValue('N'.$i,$row['kec']);
    $sheet->setCellValue('O'.$i,$row['kodepos']);
    $sheet->setCellValue('P'.$i,$row['ttinggal']);
    $sheet->setCellValue('Q'.$i,$row['transport']);
    $sheet->setCellValue('R'.$i,$row['nohp']);
    $sheet->setCellValue('S'.$i,$row['notelp']);
    $sheet->setCellValue('T'.$i,$row['email']);
    $sheet->setCellValue('U'.$i,$row['kpspkh']);
    $sheet->setCellValue('V'.$i,$row['nokpspkh']);
    $sheet->setCellValue('W'.$i,$row['kwn']);
    $sheet->setCellValue('X'.$i,$row['namaneg']);
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
$sheet->getStyle('A1:X'.$i)->applyFromArray($styleArray);

$writer         = new Xlsx($spreadsheet);
$writer->save('Report Data Pribadi.Xlsx');
if ($writer) { 
    header("location: prosesdatadiri.php"); 
} 
?>