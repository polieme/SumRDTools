using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SumRDTools
{
    public class ExcelUtils
    {
        //获取单元格的值
        public static string getCellValueByCellType(ISheet sheet, int rowIndex, int colIndex)
        {

            //获取到单元格
            ICell cell = sheet.GetRow(rowIndex).GetCell(colIndex);
            // 判断单元格是否存在  
            if (cell != null)
            {
                // 判断单元格类型  
                if (cell.CellType == CellType.String)
                {
                    // 字符串类型  
                    return cell.StringCellValue;
                }
                else if (cell.CellType == CellType.Numeric)
                {
                    // 数字类型  
                    return cell.NumericCellValue.ToString();
                }
                else if (cell.CellType == CellType.Formula)
                {
                    // 公式类型，需要计算后获取值  
                    return "" + cell.NumericCellValue;
                }
                else if (cell.CellType == CellType.Blank)
                {
                    // 空白类型，没有值  
                    return "";
                }
                else if (cell.CellType == CellType.Error)
                {
                    // 错误类型，需要处理错误情况  
                    return ""; // 需要根据实际情况处理错误值  
                }
                else
                {
                    //TODO 扔异常，父级拿到异常后，加入到错误日志中
                    return "";
                }
            }
            else
            {
                return "";
            }
        }

        //写值到单元格中
        public static void writeDataIntoCell(ISheet sheet, int rowIndex, int colIndex, dynamic cellValue)
        {
            sheet.GetRow(rowIndex).GetCell(colIndex).SetCellValue("" + cellValue);
        }
    }
}
