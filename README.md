# ExcelHandler
##Class-ExcelHandler
/**  
 *ExcelHandler  
 *对PHPExcel中读取的excel文件文件的简单封装  
 *  
 *@author lucifer  
 *@date 2014/04/08  
 *@lastmodify 2015/10/31  
 *  
 *@member  object  m_phpExcel		当前锁读取的excel文件对象  
 *@member  object  m_activeSheet	当前激活页	  
 */
##public ExcelHandler::__construct()
/**  
*根据给定路径构造ExcelHandler对象  
*初始化m_phpExcel和m_activeSheet  
*  
*@name ExcelHandler::__construct( $_filePath )  
*@param string 待解析的excel文件的路径  
*/  
##public ExcelHandler::getValueByPos()
/**  
*获取活动工作表中指定坐标下的值(按人类习惯，从1行1列开始)  
*适配PHPExcel::getCellByColumnAndRow(col, row)方法  
*  
*@name Excel::getValueByPos( $_row, $_col )  
*@param int _row  待获取的值所在的行(下标从1开始)  
*@param int _col	待获取的值在所在的列(下标从1开始)  
*/  
##public ExcelHandler::getMatrixByArea()
/**  
*获取活动工作表中指定区域的值，构建并返回一个矩阵  
*  
*@name ExcelHandler::getMatrixByArea  
*@param int _startRow  开始行,下标从1开始  
*@param int  _startCol  开始行,下标从1开始  
*@param int  _rowNum 总共要读取的行数(包含起始行)  
*@param int  _colNum 总共要读取的列数(包含起始行)  
*@return 含有指定区域信息的二维数组  
*/  
##public ExcelHandler::fshowMatrix()
/**  
*格式化显示获取的矩阵  
*将给定的矩阵转换成HTML表格形式的文档  
*  
*@name ExcelHandle::fshowMatrix( $_matrix )  
*@param array _matrix 待显示的二维数组  
*@return 格式化后的HTML字符串  
*/  
##public ExcelHandler::_get()
##public ExcelHandler::_set()