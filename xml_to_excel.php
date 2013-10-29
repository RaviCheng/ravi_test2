<?php
/**
 * xml to excel
 */

$_HallId = 6;

// db link
include_once("db.php");

// XmlExcelExport class
include_once('class/XmlExcelExport.php');

$objPHPExcel = new XmlExcelExport();

// output檔名
$objPHPExcel->generateXMLHeader("Esball");

$sql = "SELECT `LevelId`,`Script` FROM `TransferLimitByHall` WHERE `HallId`='$_HallId' ORDER BY `LevelId` ASC";
$result = mysql_query($sql) or die('MySQL query error');

while ($row = mysql_fetch_array($result)) {

    // 分頁標籤
    $objPHPExcel->worksheetStart($row['Script']);

    $LevelId = $row['LevelId'];

    // 分頁內容(抓取該分層使用者帳號)
    if ($LevelId == '0') {

        $sql = "SELECT `M`.`USERNAME` AS `USERNAME`
                    FROM `MEMBERS_Durian` AS `M`
                    LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                    WHERE `M`.`HALLID` = '{$_HallId}' AND `M`.`Id` NOT IN (SELECT `UserId` FROM `TransferUserLevelList` WHERE `LevelId` > 0)
                    ORDER BY `M`.`USERNAME`
                    LIMIT 20000
                   ";
    } else {
        $sql = "SELECT  `M`.`USERNAME` AS `USERNAME`
                  FROM `MEMBERS_Durian` AS `M`
                  LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                  WHERE `M`.`HALLID` = '{$_HallId}' AND `L`.`LevelId` = '{$LevelId}'
                  ORDER BY `M`.`USERNAME`
                  LIMIT 20000
                  ";
    }

    $result2 = mysql_query($sql) or die('MySQL query error');

    $userArray = array();
    while ($row2 = mysql_fetch_array($result2)) {
        array_push($userArray, array($row2['USERNAME']));
    }

    $objPHPExcel->setTableRows($userArray);
    // 結束分頁
    $objPHPExcel->worksheetEnd();

}

// xml結尾
$objPHPExcel->generateXMLFoot();
exit();
