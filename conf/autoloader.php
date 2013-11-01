<?
/**
 * web autoload setting
 * Ravi
 * 2013-10-30
 */

include_once('config.php');

// 目前專案所需要的 include_path
$include_path[] = APP_REAL_PATH.'/class';
//$include_path[] = APP_REAL_PATH.'/PHPExcelClasses';

set_include_path(join(PATH_SEPARATOR, $include_path));

function __autoload($className)
{
    $className = ltrim($className, '\\');
    $fileName  = '';
    $namespace = '';
    if ($lastNsPos = strripos($className, '\\')) {
        $namespace = substr($className, 0, $lastNsPos);
        $className = substr($className, $lastNsPos + 1);
        $fileName  = str_replace('\\', DIRECTORY_SEPARATOR, $namespace).DIRECTORY_SEPARATOR;
    }
    $fileName .= str_replace('_', DIRECTORY_SEPARATOR, $className).'.php';

    require $fileName;
}