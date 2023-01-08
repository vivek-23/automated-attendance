<?php

require_once __DIR__ . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->mergeCells('A2:D2');
$sheet->setCellValue('A2', 'Month - '. date("F"). " " . date("Y"));
$sheet->mergeCells('A3:D3');
$sheet->setCellValue('A3', 'Employee Name : Vivek Tadpatri');

$sheet->setCellValue('A4', 'Date');
$sheet->setCellValue('B4', 'Day');
$sheet->setCellValue('C4', 'In - Time');
$sheet->setCellValue('D4', 'Out - Time');

$curr = date("Y-m-"). "01";
$end = date("Y-m-d", strtotime(date("Y-m-t"). " +1 days"));

$row = 5;

for($i = 1;$curr != $end ; ++$i, $row++){
    $curr = strtotime($curr);
    $sheet->setCellValue('A' . $row, date("d-M-y", $curr));
    $sheet->setCellValue('B' . $row, date("l",$curr));
    if(in_array(date("w", $curr), ["6","0"])){
        $sheet->mergeCells('C' . $row . ':D'. $row);
        $sheet->setCellValue('C' . $row, 'WEEK OFF');
    }else{
        $time = getRandomWorkHours(date("Y-m-d", $curr));
        $sheet->setCellValue('C' . $row, $time[0]->format('H:i:s'));
        $sheet->setCellValue('D' . $row, $time[1]->format('H:i:s'));
    }
    $curr = date("Y-m-d", strtotime(date("Y-m-d", $curr). " +1 days"));
}

$writer = new Xlsx($spreadsheet);
$writer->save(date("F"). " " . date("Y"). '.xlsx');

function getRandomWorkHours($date){
    $tStarts = [strtotime($date . " 09:00:00"), strtotime($date . "11:00:00")];

    $tStart = rand($tStarts[0], $tStarts[1]);

    $start = new DateTimeImmutable($date);
    $start = $start->setTime(date("H", $tStart), date("i", $tStart));
    $end = $start->add(new DateInterval('PT9H'. rand(1, 50). 'M'));
    return [$start, $end];
}