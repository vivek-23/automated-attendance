<?php

require_once __DIR__ . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;

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

// make text center aligned and bold in the specified cells
foreach(['A2:D2','A3:D3','A4','B4','C4','D4'] as $cells){
   $sheet->getStyle($cells)->applyFromArray(getCenterStyle());
   $sheet->getStyle($cells)->getFont()->setBold(true);
}
// making text center aligned and bold in the specified cells ends

$curr = date("Y-m-"). "01";
$end = date("Y-m-d", strtotime(date("Y-m-t"). " +1 days"));

$row = 5;

for($i = 1;$curr != $end ; ++$i, $row++){
    $curr = strtotime($curr);
    $sheet->setCellValue('A' . $row, date("d-M-y", $curr));
    $sheet->getStyle('A' . $row)->applyFromArray(getCenterStyle());
    $sheet->setCellValue('B' . $row, date("l",$curr));
    $sheet->getStyle('B' . $row)->applyFromArray(getCenterStyle());

    if(in_array(date("w", $curr), ["6","0"])){
        $sheet->getStyle('A'. $row .':D'. $row)->getFill()->applyFromArray(getCellFillStyle());
        $sheet->mergeCells('C' . $row . ':D'. $row);
        $sheet->setCellValue('C' . $row, 'WEEK OFF');
        $sheet->getStyle('C' . $row . ':D'. $row)->applyFromArray(getCenterStyle());
        $sheet->getStyle('C' . $row . ':D'. $row)->getFont()->setBold(true);
    }else{
        $time = getRandomWorkHours(date("Y-m-d", $curr));
        $sheet->setCellValue('C' . $row, $time[0]->format('H:i:s'));
        $sheet->getStyle('C' . $row)->applyFromArray(getCenterStyle());
        $sheet->setCellValue('D' . $row, $time[1]->format('H:i:s'));
        $sheet->getStyle('D' . $row)->applyFromArray(getCenterStyle());
    }
    $curr = date("Y-m-d", strtotime(date("Y-m-d", $curr). " +1 days"));
}

$writer = new Xlsx($spreadsheet);
$fileName = date("F"). " " . date("Y"). '.xlsx';
$writer->save($fileName);

$mail = new PHPMailer(true);

try {
    //Server settings
    $mail->isSMTP();
    $mail->Host       = 'smtp.gmail.com';  
    $mail->SMTPAuth   = true;  
    $mail->Username   = 'vivekpisces23@gmail.com'; 
    $mail->Password   = json_decode(file_get_contents('password.txt'), true)['password']; 
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS; 
    $mail->Port       = 587;

    $mail->setFrom('vivekpisces23@gmail.com', 'Automated Attendance Bot Notifier');
    $mail->addAddress('vivekpisces23@gmail.com', 'Software Engineer');
    $mail->addAttachment(__DIR__ . DIRECTORY_SEPARATOR . $fileName);
    $mail->isHTML(true);
    $mail->Subject = date("F"). " ". date("Y"). " Attendance";
    $mail->Body    = '<p>Hi,</p><p>Attached is your '. date("F"). " ". date("Y"). " timesheet!</p>";
    $mail->send();
    echo "Mail sent successfully!";
    unlink(__DIR__ . DIRECTORY_SEPARATOR . $fileName);
} catch (Exception $e) {
    echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
}