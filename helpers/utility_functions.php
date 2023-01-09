<?php

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

function getRandomWorkHours($date){
    $tStarts = [strtotime($date . " 09:00:00"), strtotime($date . "11:00:00")];

    $tStart = rand($tStarts[0], $tStarts[1]);

    $start = new DateTimeImmutable($date);
    $start = $start->setTime(date("H", $tStart), date("i", $tStart));
    $end = $start->add(new DateInterval('PT9H'. rand(1, 50). 'M'));
    return [$start, $end];
}

function getCenterStyle(){
    return array(
        'alignment' => array(
            'horizontal' => Alignment::HORIZONTAL_CENTER,
        )
    );    
}

function getCellFillStyle($color = 'BABABA'){
    return array(
        'fillType' => Fill::FILL_SOLID,
        'startColor' => array('rgb' => $color)
    );
}