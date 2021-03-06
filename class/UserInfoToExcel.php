<?php
/**
 * 層級帳號輸出excel檔
 */

class UserInfoToExcel
{
    protected $_HallId = 6;
    protected $_db;

    public function __construct()
    {
        $this->_db = new DbAccess();
    }

    /**
     * 輸出xml格式的excel檔
     */
    public function CreateXmlToExcel()
    {
        // excel文件屬性設定
        $excelProp = array(
            'Author'     => 'TestA',
            'Company'    => 'TestExcel',
            'Created'    => date("Y-m-d H:i:s"),
            'Keywords'   => 'TestExcel',
            'LastAuthor' => 'TestB',
            'Version'    => '1.0.0'
        );



        $objPHPExcel = new XmlExcelExport();

        // xml標頭and檔名
        $objPHPExcel->generateXMLHeader(date('YmdHis'));

        // 給予文件屬性
        $objPHPExcel->setDocProp($excelProp);

        // 所有分層id
        $LevelId = $this->GetLevelId();

        foreach ($LevelId as $val) {
            // 工作表標籤
            $objPHPExcel->worksheetStart($val['Script']);

            // 分層內的username
            $UserInfo = $this->GetLevelUser($val['LevelId']);

            // 寫入資料列
            $objPHPExcel->setTableRows($UserInfo);

            // 結束工作表
            $objPHPExcel->worksheetEnd();
        }

        // xml結尾
        $objPHPExcel->generateXMLFoot();
    }

    /**
     * 取得該廳內的所有分層id
     *
     * @return mixed
     */
    private function GetLevelId()
    {
        $sql = "SELECT `LevelId`,`Script`
                  FROM `TransferLimitByHall`
                 WHERE `HallId`='{$this->_HallId}'
              ORDER BY `LevelId` ASC";

        $this->_db->query($sql);

        $LevelInfo = array();

        while ($row = $this->_db->fetchArray()) {
            $LevelInfo[] = array(
                'LevelId' => $row['LevelId'],
                'Script'  => $row['Script']
            );
        }

        return $LevelInfo;
    }

    /**
     * 取得該分層內所有所有username
     *
     * @param $LevelId 分層id
     * @return mixed
     */
    private function GetLevelUser($LevelId)
    {
        // 分頁內容(抓取該分層使用者帳號)
        if ($LevelId == '0') {
            $sql = "SELECT `M`.`Id` AS `UserId` , `M`.`USERNAME` AS `USERNAME`
                      FROM `MEMBERS_Durian` AS `M`
                LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                    WHERE `M`.`HALLID` = '{$this->_HallId}' AND `M`.`Id` NOT IN (SELECT `UserId` FROM `TransferUserLevelList` WHERE `LevelId` > 0)
                 ORDER BY `M`.`USERNAME`
                ";
        } else {
            $sql = "SELECT `M`.`Id` AS `UserId` , `M`.`USERNAME` AS `USERNAME`
                      FROM `MEMBERS_Durian` AS `M`
                 LEFT JOIN `TransferUserLevelList` AS `L` ON `M`.`ID` = `L`.`UserId`
                     WHERE `M`.`HALLID` = '{$this->_HallId}' AND `L`.`LevelId` = '{$LevelId}'
                  ORDER BY `M`.`USERNAME`
                 ";
        }

        $this->_db->query($sql);

        $UserInfo = array();

        // 組合成組合成xml轉excel能解析的array:
        /*
             array(
                array('a'),     // 第一列
                array('b'),     // 第二列
                array('c', 'd') // 第三列
             );
        */
        while ($row = $this->_db->fetchArray()) {
            $UserInfo[] = array(
                preg_match('/^[0-9]*$/', $row['USERNAME']) ? '*'.($row['USERNAME']) : $row['USERNAME']
            );
        }

        return $UserInfo;
    }
}