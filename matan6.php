<?php
/**
 * Created by PhpStorm.
 * User: ruartel
 * Date: 2/28/16
 * Time: 11:37 PM
 */
/** Include PHPExcel */
require_once dirname(__FILE__) . '/meekrodb.matan.php';
require_once dirname(__FILE__) . '/phpExcel/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();
$objPHPExcel->getProperties()->setCreator("Gla Solutions");

$objPHPExcel->setActiveSheetIndex(0);
$cols_result = DB::query('SHOW COLUMNS FROM project_tafkid');
$x='A';
$colsName = array();
foreach($cols_result as $col){
    $objPHPExcel->setActiveSheetIndex(0)->setCellValue($x . '1', $col['Field']);
    $colsName[$x] = $col['Field'];
    $x++;
}

$row=2;
$result = DB::query('SELECT * FROM project_tafkid');
foreach ($result as $k => $inst){
    $x='A';
    foreach($inst as $k=>$v){
        $value = $inst[$colsName[$x]];
        $objPHPExcel->setActiveSheetIndex(0)->setCellValue($x . $row, $value);
        $x++;
    }
    ++$row;
}

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));

//$result2 = DB::query('SELECT * from institute2area WHERE inst_id=' . $inst['id']);
//foreach($result2 as $area){
//    echo '<br />';
//    echo $areas[$area['area_id']];
//}
//$result3 = DB::query('SELECT * from institute2spec WHERE inst_id=' . $inst['id']);
//foreach($result3 as $spec){
//    echo '<br />';
//    echo $specialities[$spec['spec_id']];
//}
//$result4 = DB::query('SELECT * from institute_files WHERE inst_id=' . $inst['id']);
//foreach($result4 as $file){
//    echo '<br />';
//    var_dump($file);
//}


// Redirect output to a client’s web browser (Excel5)
//header('Content-Type: application/vnd.ms-excel');
//header('Content-Disposition: attachment;filename="matan5.xls"');
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

