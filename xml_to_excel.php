<?php
/**
 * output excel
 */

include_once("conf/autoloader.php");

$CreateExcel = new UserInfoToExcel();
$CreateExcel->CreateXmlToExcel();
