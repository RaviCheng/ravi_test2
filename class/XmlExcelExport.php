<?php
/**
 * Class XmlExcelExport
 * xml to excel
 */
class XmlExcelExport
{

    /**
     * 文檔頭標籤
     * @var string
     */
    private $header = "<?xml version=\"1.0\" encoding=\"%s\"?\>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x =\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org /TR/REC-html40\">";

    /**
     * 文檔尾標籤
     * @var string
     */
    private $footer = "</Workbook>";

    /**
     * 內容編碼
     * @var string
     */
    private $sEncoding;

    /**
     * 是否轉換特定字段值的類型
     * @var boolean
     */
    private $bConvertTypes;

    /**
     * 生成的Excel內工作表個數
     *
     * @var int
     */
    private $dWorksheetCount = 0;

    /**
     * 構造函數
     * 使用類型轉換時要確保:頁碼和郵編號以'0'開頭
     * @param string  $sEncoding     內容編碼
     * @param boolean $bConvertTypes 是否轉換特定字段值的類型
     */
    function __construct($sEncoding = 'UTF-8', $bConvertTypes = false)
    {
        $this->bConvertTypes = $bConvertTypes;
        $this->sEncoding     = $sEncoding;
    }

    /**
     * 返回工作簿標題,最大字符數為31
     * @param string $title 工作簿標題
     * @return string
     */
    function getWorksheetTitle($title = 'Table1')
    {
        $title = preg_replace(
            "/[\\\|:|\/|\?|\*|\[|\]]/",
            "",
            empty ($title) ? 'Table'.($this->dWorksheetCount + 1) : $title
        );
        return substr($title, 0, 31);
    }

    /**
     * 向客戶端發送Excel頭信息
     *
     * @param string $filename 文件名稱,不能是中文
     */
    function generateXMLHeader($filename)
    {

//        $filename = preg_replace('/[^aA-zZ0-9\_\-]/', '', $filename);
//        $filename = urlencode($filename);

        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Content-Type: application/force-download");
        header("Content-Type: application/vnd.ms-excel; charset={$this->sEncoding}");
        header("Content-Transfer-Encoding: binary");
        header("Content-Disposition: attachment; filename={$filename}.xls");

        echo stripslashes(sprintf($this->header, $this->sEncoding));
    }

    /**
     * 向客戶端發送Excel結束標籤
     */
    function generateXMLFoot()
    {
        echo $this->footer;
    }

    /**
     * 開啟工作簿
     * @param string $title
     */
    function worksheetStart($title)
    {
        $this->dWorksheetCount ++;
        echo "\n<Worksheet ss:Name=\"".$this->getWorksheetTitle($title)."\">\n<Table>\n";
    }

    /**
     * 結束工作簿
     */
    function worksheetEnd()
    {
        echo "</Table>\n</Worksheet>\n";
    }

    /**
     * 設置表頭信息
     * @param array $header
     */
    function setTableHeader(array $header)
    {
        echo $this->_parseRow($header);
    }

    /**
     * 設置表內行記錄數據
     * @param array $rows 多行記錄
     */
    function setTableRows(array $rows)
    {
        foreach ($rows as $row) {
            echo $this->_parseRow($row);
        }
    }

    /**
     * 將傳人的單行記錄數組轉換成xml 標籤形式
     */
    private function _parseRow(array $row)
    {
        $cells = "";
        foreach ($row as $k => $v) {
            $type = 'String';
            if ($this->bConvertTypes === true && is_numeric($v)) {
                $type = 'Number';
            }

            $v = htmlentities($v, ENT_COMPAT, $this->sEncoding);
            $cells .= "<Cell><Data ss:Type=\"$type\">".$v."</Data></Cell>";
        }
        return "<Row>".$cells."</Row>\n";
    }

}

/**
 * 解決發出的文件名稱中文亂碼的問題
 * @param        $serverPath
 * @param        $filename
 * @param string $charset
 * @param string $mimeType
 */
function sendFile($serverPath, $filename, $charset = 'UTF-8', $mimeType = 'application/octet-stream')
{
    // 文件名亂碼問題
    if (preg_match("/MSIE/", $_SERVER ["HTTP_USER_AGENT"])) {
        $filename         = urlencode($filename);
        $filename         = str_replace("+", "%20", $filename); //替換空格
        $attachmentHeader = "Content-Disposition: attachment; filename=\"{$filename}\"; charset={$charset}";
    } else {
        if (preg_match("/Firefox/", $_SERVER ["HTTP_USER_AGENT"])) {
            $attachmentHeader = 'Content-Disposition: attachment; filename*="utf8\'\''.$filename.'"';
        } else {
            $attachmentHeader = "Content-Disposition: attachment; filename=\"{$filename}\"; charset={$charset}";
        }
    }


    $filesize = filesize($serverPath);

    //header("Pragma: public"); header("Expires: 0");
    header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
    header("Content-Type: application/force-download");
    header("Content-Type: {$mimeType}");

    header($attachmentHeader);
    header('Pragma: cache');
    header('Cache-Control: public, must-revalidate, max-age=0');
    header("Content-Length: {$filesize}");
    readfile($serverPath);
    exit;
}

//$data = array(
//    array( 'a','b','','d'),array( 'e','f','','h')
//);

//$excel = new XmlExcelExport();
//$excel->generateXMLHeader("測試");
//$excel->worksheetStart('s1');
//$excel->setTableRows($data);
//$excel->worksheetEnd();
//$excel->worksheetStart('s2');
//$excel->setTableRows($data);
//$excel->worksheetEnd();
//
//$excel->generateXMLFoot();