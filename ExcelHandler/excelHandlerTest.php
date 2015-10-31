<?php
/** excelHandler.class.php的参试文件 **/
header("content-type:text/html; charset=utf-8");
include("PHPExcel/PHPExcel.php");
include("excelHandler.class.php");

$filePath = "priceSheets.xls";
$excelH = new ExcelHandler( $filePath );

/** 获取1行1列(1, A) 的值
echo "获取1行1列(1, A) 的值: ";
echo $excelH->getValueByPos(1, 1), "<br />";
echo "获取2行3列(2, C)的值: ";
echo $excelH->getValueByPos(2, 3), "<br />";  **/

/** 获取(27, B) 到  (36, L)的值 B=2, L=12
并格式化显示
**/
$matrix = $excelH->getMatrixByArea(27, 2, 10, 11);
$htmlStr = $excelH->fshowMatrix( $matrix );

echo $htmlStr;
