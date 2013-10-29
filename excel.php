<?php
// post
$postLesson  = $_POST['txtLesson'];
$postTeacher = $_POST['txtTeacher'];
$postPrice   = $_POST['txtPrice'];

// Include PHPExcel
require_once 'PHPExcelClasses/PHPExcel.php';

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

$objPHPExcel->setActiveSheetIndex(0);

// 標籤名稱設置
$objPHPExcel->getActiveSheet()->setTitle("課程資料");

// 合併儲存格
$objPHPExcel->getActiveSheet()->mergeCells('B1:H1');
$objPHPExcel->getActiveSheet()->mergeCells('B2:H2');
$objPHPExcel->getActiveSheet()->mergeCells('B3:H3');
$objPHPExcel->getActiveSheet()->mergeCells('B4:H4');
$objPHPExcel->getActiveSheet()->mergeCells('B5:H5');

// 設置列寬度
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(20);

// 字串格式
$StrType = PHPExcel_Cell_DataType::TYPE_STRING;

// 標題
$objPHPExcel->getActiveSheet()->setCellValue('A1', '課程名稱:');
$objPHPExcel->getActiveSheet()->setCellValue('A2', '授課教師:');
$objPHPExcel->getActiveSheet()->setCellValue('A3', '上課日期時間:');
$objPHPExcel->getActiveSheet()->setCellValue('A4', '報名截止日:');
$objPHPExcel->getActiveSheet()->setCellValue('A5', '報名費用:');

// 課程資料
$objPHPExcel->getActiveSheet()->setCellValueExplicit('B1', $postLesson, $StrType);
$objPHPExcel->getActiveSheet()->setCellValueExplicit('B2', $postTeacher, $StrType);
$objPHPExcel->getActiveSheet()->setCellValueExplicit('B3', date("Y-m-d H:i:s"), $StrType);
$objPHPExcel->getActiveSheet()->setCellValueExplicit('B4', date("Y-m-d H:i:s"), $StrType);
$objPHPExcel->getActiveSheet()->setCellValueExplicit('B5', $postPrice, $StrType);

// 設置對齊
$objPHPExcel->getActiveSheet()->getStyle('A1:A8')->getAlignment()->setHorizontal(
    PHPExcel_Style_Alignment::HORIZONTAL_RIGHT
);
$objPHPExcel->getActiveSheet()->getStyle('B1:B8')->getAlignment()->setHorizontal(
    PHPExcel_Style_Alignment::HORIZONTAL_LEFT
);
$objPHPExcel->getActiveSheet()->getStyle('B1:B8')->getAlignment()->setVertical(
    PHPExcel_Style_Alignment::VERTICAL_TOP
);


/* 報名資料 - 新分頁*/
$objPHPExcel->createSheet();

$objPHPExcel->setActiveSheetIndex(1);

// 標籤名稱設置
$objPHPExcel->getActiveSheet()->setTitle("報名資料");

//標題
$objPHPExcel->getActiveSheet()->setCellValue('B1', '姓名');
$objPHPExcel->getActiveSheet()->setCellValue('C1', '生日');
$objPHPExcel->getActiveSheet()->setCellValue('D1', '電話');
$objPHPExcel->getActiveSheet()->setCellValue('E1', '信箱');
$objPHPExcel->getActiveSheet()->setCellValue('F1', '地址');

// 設置列寬度
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(15);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(18);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(30);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(50);

// 人數計數
$studCount = 0;

// 撈取學生資料
for ($i = 1; $i <= 10000; $i ++) {
    $studCount ++;
    $objPHPExcel->getActiveSheet()->setCellValueExplicit('A'.($studCount + 1), $studCount, $StrType);
    $objPHPExcel->getActiveSheet()->setCellValueExplicit('B'.($studCount + 1), "郭小猴", $StrType);
    $objPHPExcel->getActiveSheet()->setCellValueExplicit('C'.($studCount + 1), "100/09/08", $StrType);
    $objPHPExcel->getActiveSheet()->setCellValueExplicit('D'.($studCount + 1), "0912345678", $StrType);
    $objPHPExcel->getActiveSheet()->setCellValueExplicit('E'.($studCount + 1), "test@gmail.com", $StrType);
    $objPHPExcel->getActiveSheet()->setCellValueExplicit('F'.($studCount + 1), "台中市南屯區9999號", $StrType);
}

// 學生資料統計資訊
$objPHPExcel->getActiveSheet()->mergeCells('A'.($studCount + 3).':F'.($studCount + 3));
$objPHPExcel->getActiveSheet()->setCellValueExplicit('A'.($studCount + 3), "共計 ".($studCount)." 員", $StrType);
$objPHPExcel->getActiveSheet()->getStyle('A'.($studCount + 3))->getAlignment()->setHorizontal(
    PHPExcel_Style_Alignment::HORIZONTAL_CENTER
);

// 設置對齊
$objPHPExcel->getActiveSheet()->getStyle('B1:F1')->getAlignment()->setHorizontal(
    PHPExcel_Style_Alignment::HORIZONTAL_LEFT
);
$objPHPExcel->getActiveSheet()->getStyle('A2:A'.($studCount + 1))->getAlignment()->setHorizontal(
    PHPExcel_Style_Alignment::HORIZONTAL_RIGHT
);


// Redirect output to a client’s web browser (Excel5,2003格式)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="php 初階班.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');


$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;
