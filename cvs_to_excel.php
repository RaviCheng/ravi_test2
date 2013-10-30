<?php
/**
 * phpexcel to excel file
 * ps. 尚有緩衝不足問題(超過15000筆)
 */

$_HallId = 6;

// db link
include_once("class/DbAccess.php");

// phpexcel class
include_once('PHPExcelClasses/PHPExcel.php');

$objPHPExcel = new PHPExcel();

$sql = "SELECT `LevelId`,`Script` FROM `TransferLimitByHall` WHERE `HallId`='$_HallId' ORDER BY `LevelId` ASC";
$result = mysql_query($sql) or die('MySQL query error');


$levelCount = 0;

while ($row = mysql_fetch_array($result)) {

    $objPHPExcel->setActiveSheetIndex($levelCount);

    // 分頁標籤
    $objPHPExcel->getActiveSheet()->setTitle($row['Script']);

    $LevelId = $row['LevelId'];

    // 分頁內容(抓取該分層使用者帳號)
    if ($LevelId == '0') {

        $sql = "SELECT `M`.`Id` AS `UserId` , `M`.`USERNAME` AS `USERNAME`
                    FROM `MEMBERS_Durian` AS `M`
                    LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                    WHERE `M`.`HALLID` = '{$_HallId}' AND `M`.`Id` NOT IN (SELECT `UserId` FROM `TransferUserLevelList` WHERE `LevelId` > 0)
                    ORDER BY `M`.`USERNAME`
                    LIMIT 10000";
    } else {
        $sql = "SELECT `M`.`Id` AS `UserId` , `M`.`USERNAME` AS `USERNAME`
                  FROM `MEMBERS_Durian` AS `M`
                  LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                  WHERE `M`.`HALLID` = '{$_HallId}' AND `L`.`LevelId` = '{$LevelId}'
                  ORDER BY `M`.`USERNAME`
                  LIMIT 10000";
    }

    $result2 = mysql_query($sql) or die('MySQL query error');

    $userCount = 0;
    while ($row2 = mysql_fetch_array($result2)) {
//        $objPHPExcel->getActiveSheet()->setCellValueExplicit(
//            'A'.($userCount + 1),
//            ($userCount + 1),
//            PHPExcel_Cell_DataType::TYPE_STRING
//        );

        $objPHPExcel->getActiveSheet()->setCellValueExplicit(
            'A'.($userCount + 1),
            $row2['USERNAME'],
            PHPExcel_Cell_DataType::TYPE_STRING
        );

        $userCount ++;
    }

    // 產生下一張分頁
    $objPHPExcel->createSheet();
    $levelCount ++;
}


//header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//header('Content-Disposition: attachment;filename="2007test.xlsx"');
//header('Cache-Control: max-age=0');
//
header ( "Content-Type: text/csv" );
header ( "Content-Disposition:filename=csv_test.csv" );
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV');
$objWriter->save('php://output');
// Save Excel 2007 file
//echo date('H:i:s'), " Write to Excel2007 format";
//echo '<br>';
//$callStartTime = microtime(true);

//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
//$objWriter->save('file/2007test.xlsx');
//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
//$objWriter->save('file/2007test.xls');

//$callEndTime = microtime(true);
//$callTime    = $callEndTime - $callStartTime;
//
//echo date('H:i:s'), " File written to 2007test.xlsx";
//echo '<br>';
//echo 'Call time to write Workbook was ', sprintf('%.4f', $callTime), " seconds";
//echo '<br>';
//echo date('H:i:s'), ' Current memory usage: ', (memory_get_usage(true) / 1024 / 1024), " MB";
//echo '<br>';
exit();
