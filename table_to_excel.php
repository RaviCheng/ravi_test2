<?php
/**
 * html table to excel
 */

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="2007test.xlsx"');
header('Cache-Control: max-age=0');

$_HallId = 6;

// db link
include_once("class/DbAccess.php");

// TableExcelExport class
include_once('class/TableExcelExport.php');
$excel = new TableExcelExport();

// 設定第一行格式(a:string , 0:number)
$excel->colsAttrib('a');


$sql = "SELECT `LevelId`,`Script` FROM `TransferLimitByHall` WHERE `HallId`='$_HallId' ORDER BY `LevelId` ASC";
$result = mysql_query($sql) or die('MySQL query error');


$levelCount = 0;

while ($row = mysql_fetch_array($result)) {

    $LevelId = $row['LevelId'];

    // 分頁內容(抓取該分層使用者帳號)
    if ($LevelId == '0') {

        $sql = "SELECT `M`.`Id` AS `UserId` , `M`.`USERNAME` AS `USERNAME`
                    FROM `MEMBERS_Durian` AS `M`
                    LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                    WHERE `M`.`HALLID` = '{$_HallId}' AND `M`.`Id` NOT IN (SELECT `UserId` FROM `TransferUserLevelList` WHERE `LevelId` > 0)
                    ORDER BY `M`.`USERNAME`
                    LIMIT 20000";
    } else {
        $sql = "SELECT `M`.`Id` AS `UserId` , `M`.`USERNAME` AS `USERNAME`
                  FROM `MEMBERS_Durian` AS `M`
                  LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                  WHERE `M`.`HALLID` = '{$_HallId}' AND `L`.`LevelId` = '{$LevelId}'
                  ORDER BY `M`.`USERNAME`
                  LIMIT 20000";
    }

    $result2 = mysql_query($sql) or die('MySQL query error');

    $userCount = 0;
    while ($row2 = mysql_fetch_array($result2)) {
        $excel->excelWrite(array($row2['USERNAME']));
        $userCount ++;
    }

    $levelCount ++;
}

$excel->excelEnd();
exit();
