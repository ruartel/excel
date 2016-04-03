<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/meekrodb.mwall.php';
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
$arrNames[1]='fname';
$arrNames[2]='lname';
$arrNames[3]='email';
$arrNames[4]='vdate';
$arrNames[5]='payment_amount';
$arrNames[6]='pageFrom';
$arrNames[7]='wallName1';
$arrNames[8]='wallName2';
$arrNames[9]='wallName3';
$arrNames[10]='payment_method';
$arrNames[11]='address1';
$arrNames[12]='address2';
$arrNames[13]='city';
$arrNames[14]='state';
$arrNames[15]='zip';
$arrNames[16]='country';
$arrNames[17]='donation_type';

$arrToExcel=array();
$arrToExcel[0]='B';
$arrToExcel[1]='C';
$arrToExcel[2]='D';
$arrToExcel[3]='E';
$arrToExcel[4]='F';
$arrToExcel[5]='G';
$arrToExcel[6]='H';
$arrToExcel[7]='I';
$arrToExcel[8]='J';
$arrToExcel[9]='K';
$arrToExcel[10]='L';
$arrToExcel[11]='M';
$arrToExcel[12]='N';
$arrToExcel[13]='O';
$arrToExcel[14]='P';
$arrToExcel[15]='Q';
$arrToExcel[16]='R';
$arrToExcel[17]='S';

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
            ->setCellValue('R1', 'Donor Country')
            ->setCellValue('S1', 'Donate to');


$row = 2;
$result = DB::query('select dp.id,DATE_FORMAT(cdate, "%d/%m/%Y") as vdate, dn.fname,dn.lname,dn.email,'
        . 'dn.suggested_donation as payment_amount,dp.first_name as wallName1,dp.last_name as wallName2, dp.hname as wallName3,dn.payment_method,'
        . 'dn.address1,dn.address2, dn.city, dn.state, dn.zip, dn.country, dn.donation_type'
        . ' FROM dpeople dp left join donor dn on dn.id=dp.donor_id WHERE dp.isRemoved=0 order by cdate DESC');
//var_dump($result);
foreach ($result as $k => $u){
//    $arr[$u['id']]=array();
//    $arr[$u['id']][28] = $u['date_created'];
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue('A' . $row, $row);
    
    foreach ($arrNames as $col => $col_name){
        $cur_col = $arrToExcel[$col];
        if($col == 0){
            $cur_value = 'w' . $u[$col_name];
        }else{
            $cur_value = $u[$col_name];
        }
        
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($cur_col . $row, $cur_value);
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