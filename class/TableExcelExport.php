<?php
/**
 * Class SimpleExcel
 * html table to excel
 * ps. 無法生成分頁檔
 */

class TableExcelExport
{
    var $rowsNum = 0;

    var $attrib = array();

    var $in_charset = 'UTF-8';

    function SimpleExcel()
    {
        echo pack("ssssss", 0x809, 0x8, 0x0, 0x10, 0x0, 0x0);
        return;
    }

    function __construct($inCharset = '')
    {
        if (! empty ($inCharset)) {
            $this->in_charset = $inCharset;
        }

        $this->SimpleExcel();
    }

    function excelItem($string = array())
    {
        for ($i = 0; $i < count($string); $i ++) {
            $curStr = $string [$i];
            $curStr = $this->iconvToData($curStr);

            $L = strlen($curStr);
            echo pack("ssssss", 0x204, 8 + $L, $this->rowsNum, $i, 0x0, $L);
            echo $curStr;
        }
        $this->rowsNum ++;
        return;
    }

    function colsAttrib($string = array())
    {
        $this->attrib = $string;
        return;
    }

    function excelWrite($string = array())
    {
        for ($i = 0; $i < count($string); $i ++) {
            $curStr = $string [$i];
            $curStr = $this->iconvToData($curStr);

            if ($this->attrib [$i] == "1") {
                echo pack("sssss", 0x203, 14, $this->rowsNum, $i, 0x0);
                echo pack("d", $curStr);
            } else {
                $L = strlen($curStr);
                echo pack("ssssss", 0x204, 8 + $L, $this->rowsNum, $i, 0x0, $L);
                echo $curStr;
            }
        }

        $this->rowsNum ++;
    }

    function excelEnd()
    {
        echo pack("ss", 0x0A, 0x00);
        return;
    }

    function iconvToData($data)
    {
//        return iconv($this->in_charset, 'utf-8', $data);
        return $data;
    }
}

//header("Content-Type: application/vnd.ms-excel");
//header("Content-Disposition: attachment; filename=tableToExcel.xls");
//header("Pragma: no-cache");
//header("Expires: 0");
//
//$excel = new SimpleExcel(); //調用類開始
////定義屬性，數字型為"1"，字符型為"a"
//$excel->colsAttrib(array("1", "a", "a", "a", "a", "1", "a"));
//$excel->excelWrite(array("01010", "01010", "ac", "ad", "ae", "af", "ag"));
//$excel->excelWrite(array("02020", "02020", "bc", "bd", "be", "bf", "bg"));
//$excel->excelEnd();
