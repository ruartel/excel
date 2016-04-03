<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/meekrodb.fjc.php';
require_once dirname(__FILE__) . '/phpExcel/Classes/PHPExcel.php';

function getNameFromNumber($num) {
    $numeric = ($num - 1) % 26;
    $letter = chr(65 + $numeric);
    $num2 = intval(($num - 1) / 26);
    if ($num2 > 0) {
        return getNameFromNumber($num2) . $letter;
    } else {
        return $letter;
    }
}

$arrNames=array();
$arrNames[0]='id';
$arrNames[1]='firstName';
$arrNames[2]='lastName';
$arrNames[4]='email';
$arrNames[28]='vdate';
$arrNames[26]='payment_amount';
$arrNames[18]='pageFrom';
$arrNames[8]='wallName1';
$arrNames[9]='wallName2';
$arrNames[10]='wallName3';
$arrNames[32]='payment_method';
$arrNames["19.1"]='address1';
$arrNames["19.2"]='address2';
$arrNames["19.3"]='city';
$arrNames["19.4"]='state';
$arrNames["19.5"]='zip';
$arrNames["19.6"]='country';

$arrToExcel=array();
$arrToExcel[0]='B';
$arrToExcel[1]='C';
$arrToExcel[2]='D';
$arrToExcel[4]='E';
$arrToExcel[28]='F';
$arrToExcel[26]='G';
$arrToExcel[18]='H';
$arrToExcel[8]='I';
$arrToExcel[9]='J';
$arrToExcel[10]='K';
$arrToExcel[32]='L';
$arrToExcel["19.1"]='M';
$arrToExcel["19.2"]='N';
$arrToExcel["19.3"]='O';
$arrToExcel["19.4"]='P';
$arrToExcel["19.5"]='Q';
$arrToExcel["19.6"]='R';

$arr = array();
$result = DB::query('select id, DATE_FORMAT(date_created, "%d/%m/%Y") as date_created FROM wp_rg_lead WHERE form_id in (19,20)');
//var_dump($result);
foreach ($result as $k => $u){
    $arr[$u['id']]=array();
    $arr[$u['id']][28] = $u['date_created'];
    $details = DB::query('select * FROM wp_rg_lead_detail WHERE lead_id =' . $u['id']);
    foreach ($details as $d=>$l){
//        var_dump($l);
        if(isset($arrNames[$l['field_number']])){
            $arr[$u['id']][$l['field_number']] = $l['value'];
        }
    }
}

//var_dump($arr);

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Gla Solutions");

$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'stam')
            ->setCellValue('B1', 'user_ID')
            ->setCellValue('C1', 'firstName')
            ->setCellValue('D1', 'lastName')
            ->setCellValue('E1', 'email')
            ->setCellValue('F1', 'vdate')
            ->setCellValue('G1', 'payment_amount')
            ->setCellValue('H1', 'pageFrom')
            ->setCellValue('I1', 'wallName1')
            ->setCellValue('J1', 'wallName2')
            ->setCellValue('K1', 'wallName3')
            ->setCellValue('L1', 'payment_method')
            ->setCellValue('M1', 'Donor Address1')
            ->setCellValue('N1', 'Donor Address2')
            ->setCellValue('O1', 'Donor City')
            ->setCellValue('P1', 'Donor State')
            ->setCellValue('Q1', 'Donor Zip')
            ->setCellValue('R1', 'Donor Country');


$row = 2;
foreach ($arr as $k => $v){
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A' . $row, $row);
    $cur_col = $arrToExcel[0];
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cur_col . $row, $k);
    foreach ($v as $l=>$vv){
        $cur_col = $arrToExcel[$l];
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cur_col . $row, $vv);
    }
    $row++;
    
}


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));
// Redirect output to a client’s web browser (Excel5)
//header('Content-Type: application/vnd.ms-excel');
//header('Content-Disposition: attachment;filename="m_wall1.xls"');
//header('Cache-Control: max-age=0');
//// If you're serving to IE 9, then the following may be needed
//header('Cache-Control: max-age=1');
//
//// If you're serving to IE over SSL, then the following may be needed
//header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
//header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
//header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
//header ('Pragma: public'); // HTTP/1.0
//
//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save('php://output');
//exit;