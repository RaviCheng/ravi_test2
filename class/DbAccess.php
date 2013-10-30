<?php
/**
 * Created by JetBrains PhpStorm.
 * User: ravi
 * Date: 2013/10/28
 * Time: 上午 11:00
 * To change this template use File | Settings | File Templates.
 */

class DbAccess
{
    private $_dbConn;
    private $_queryResource;

    function __construct()
    {
        self::dbAccess();
    }

    private function dbAccess()
    {
        $this->_dbConn = mysql_connect(DB_HOST, DB_USER, DB_PASSWD) or die ("MySQL Connect Error");
        mysql_query("SET NAMES utf8", $this->_dbConn);
        mysql_select_db(DB_NAME, $this->_dbConn) or die ("MySQL Select DB Error");
    }

    public function dbClose()
    {
        mysql_close($this->_dbConn) or die ("MySQL close Error");;
    }

    public function query($sql)
    {
        $this->detach();
        $this->_queryResource = mysql_query($sql, $this->_dbConn) or die (mysql_error());
        return $this->_queryResource;
    }

    private function detach()
    {
        $this->_queryResource = null;
    }

    public function fetchArray()
    {
        return mysql_fetch_array($this->_queryResource, MYSQL_ASSOC);
    }

    public function fetchRow()
    {
        return mysql_fetch_row($this->_queryResource);
    }

    public function fetchNumRows()
    {
        return mysql_num_rows($this->_queryResource);
    }

    public function fetchInsertId()
    {
        return mysql_insert_id($this->_dbConn);
    }

    public function valueEscape($value)
    {
        return mysql_real_escape_string($value, $this->_dbConn);
    }
}