<?php
//ini_set("memory_limit","256M");
/** Error reporting */
//error_reporting(E_ALL);
//ini_set('display_errors', TRUE);
//ini_set('display_startup_errors', TRUE);
//date_default_timezone_set('Europe/London');

//define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

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
$arrNames[11]='name';
$arrNames[12]='last_name';
$arrNames[15]='email';
$arrNames[28]='vdate';
$arrNames[26]='payment_amount';
$arrNames[7]='cause';
$arrNames[9]='Note';
$arrNames[14]='payment_type';
$arrNames[20]='city';
$arrNames[19]='fn';
$arrNames[50]='title';
$arrNames[51]='transaction_id';
$arrNames[52]='amount';
$arrNames[53]='amount2';
$arrNames["19.1"]='address1';
$arrNames["19.2"]='address1';
$arrNames["19.3"]='city';
$arrNames["19.4"]='state';
$arrNames["19.5"]='zip';
$arrNames["19.6"]='country';
$arrNames["21"]='recurring_donation';

$arrToExcel=array();
$arrToExcel[0]='B';
$arrToExcel[11]='C';
$arrToExcel[12]='D';
$arrToExcel[15]='E';
$arrToExcel[28]='F';
$arrToExcel[26]='G';
$arrToExcel[7]='H';
$arrToExcel[9]='I';
$arrToExcel[14]='J';
$arrToExcel[20]='K';
$arrToExcel[19]='L';
$arrToExcel[50]='M';
$arrToExcel[51]='N';
$arrToExcel[52]='O';
$arrToExcel[53]='P';
$arrToExcel["19.1"]='Q';
$arrToExcel["19.2"]='R';
$arrToExcel["19.3"]='S';
$arrToExcel["19.4"]='T';
$arrToExcel["19.5"]='U';
$arrToExcel["19.6"]='V';
$arrToExcel["21"]='X';

$arr = array();
$result = DB::query('select id, DATE_FORMAT(date_created, "%d/%m/%Y") as date_created,payment_amount,transaction_id,form_id FROM wp_rg_lead WHERE payment_amount IS NOT NULL order by id desc limit 100');
//var_dump($result);
//die();
foreach ($result as $k => $u){
    if($u['id']){
    $arr[$u['id']]=array();
    $arr[$u['id']][28] = $u['date_created'];
    $arr[$u['id']][26] = $u['payment_amount'];
    $details = DB::query('select * FROM wp_rg_lead_detail WHERE value IS NOT NULL and lead_id =' . $u['id']);
    foreach ($details as $d=>$l){
        if(isset($arrNames[$l['field_number']])){
            if($l['field_number'] == 20 && ($l['value'] == 'One-Time' || $l['value'] == 'Monthly' || $l['value'] == 'Yearly')){
                $arr[$u['id']][21] = $l['value'];
            }else{
                $arr[$u['id']][$l['field_number']] = $l['value'];
            }
        }
    }
    if($u['id']){
        $details2 = DB::query('select * FROM wp_gf_addon_payment_callback WHERE lead_id =' . $u['id']);
        $details3 = DB::query('select * FROM wp_gf_addon_payment_transaction WHERE lead_id =' . $u['id']);

    }
    if($u['form_id']){
        $details4 = DB::query('select title FROM wp_rg_form WHERE id =' . $u['form_id']);
//        var_dump('4444');
//        var_dump($details4);
        $arr[$u['id']][50] = $details4[0]['title'];
    }
    
    if(isset($details2[0]['transaction_id'])){
        $arr[$u['id']][51] = $details2[0]['transaction_id'];
        $arr[$u['id']][52] = $details2[0]['amount'];
    }else if(isset($details3[0]['transaction_id'])){
        $arr[$u['id']][51] = $details3[0]['transaction_id'];
        $arr[$u['id']][53] = $details3[0]['amount'];
    }else{
        $arr[$u['id']][51] = '';
        $arr[$u['id']][52] = '';
        $arr[$u['id']][53] = '';
    }

//        var_dump('5555');
//        var_dump($arr);
    set_time_limit ( 20 );
    }
//    var_dump($arr);
//die();
}
//var_dump($arr);
//die();
$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Gla Solutions");

$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'stam')
            ->setCellValue('B1', 'user_ID')
            ->setCellValue('C1', 'name')
            ->setCellValue('D1', 'last_name')
            ->setCellValue('E1', 'email')
            ->setCellValue('F1', 'vdate')
            ->setCellValue('G1', 'payment_amount')
            ->setCellValue('H1', 'cause')
            ->setCellValue('I1', 'Note')
            ->setCellValue('J1', 'payment_type')
            ->setCellValue('K1', 'city')
            ->setCellValue('L1', 'fn')
            ->setCellValue('M1', 'title')
            ->setCellValue('N1', 'transaction_id')
            ->setCellValue('O1', 'amount')
            ->setCellValue('P1', 'amount2')
            ->setCellValue('Q1', 'Donor Address1')
            ->setCellValue('R1', 'Donor Address2')
            ->setCellValue('S1', 'Donor City')
            ->setCellValue('T1', 'Donor State')
            ->setCellValue('U1', 'Donor Zip')
            ->setCellValue('V1', 'Donor Country')
            ->setCellValue('X1', 'Recurring Donation');


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
//    set_time_limit ( 60 );
}

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));
// Redirect output to a client’s web browser (Excel5)
//header('Content-Type: application/vnd.ms-excel');
//header('Content-Disposition: attachment;filename="fjc_e.xls"');
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