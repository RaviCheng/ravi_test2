<!DOCTYPE html>
<html>
<head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>
<body>
<table>
    <tr><td align="center">課程</td></tr>
    <tr>
        <td>
            <form id="Postform" name="Postform" method="post" action="excel.php">
                <table cellpadding="0" cellspacing="0" width="100%" border="0" style="line-height:25px;">
                    <tr>
                        <td align="right" width="100">
                            課程名稱：
                        </td>
                        <td align="left">
                            <input type="text" id="txtLesson" name="txtLesson" style="border:1px solid #abadb3;width:300px;" />
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            教師姓名：
                        </td>
                        <td align="left">
                            <input type="text" name="txtTeacher" id="txtTeacher" style="border:1px solid #abadb3;width:300px;" />
                        </td>
                    </tr>
                    <tr>
                        <td align="right">
                            費用：
                        </td>
                        <td align="left">
                            <textarea name="txtPrice" id="txtPrice" style="border:1px solid #abadb3;width:450px;height:120px;"></textarea>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" style="padding:10px 0;" colspan="2" >
                            <table cellpadding="0" cellspacing="0"  border="0">
                                <tr>
                                    <td style="padding-right: 10px;">
                                        <input type="submit" name="btnSend" id="btnSend" value="產生excel檔"/>
                                    </td>
                                    <td style="padding-right: 10px;">
                                        <input type="reset" value="重新填寫" />
                                    </td>

                            </table>
                        </td>
                    </tr>
                </table>
            </form>
        </td>
    </tr>
</table>
</body>
</html>