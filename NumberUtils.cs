using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SumRDTools
{
    public class NumberUtils
    {
        //正则获取到数字返回
        public static decimal getDecimal(string numberStr)
        {

            //通过正则获取到表格中的数据
            string pattern = @"(-?\d+)(\.\d+)?"; // 匹配一串连续的数字  

            Regex regex = new Regex(pattern);
            System.Text.RegularExpressions.Match match = regex.Match(numberStr);

            if (match.Success)
            {
                if (decimal.TryParse(match.Value, out decimal result))
                {
                    return result;
                }
                else {
                    return 0;
                }
            }
            else
            {
                return 0;
            }
        }

        //正则获取到数字返回
        public static int getInt(string numberStr)
        {

            //通过正则获取到表格中的数据
            string pattern = @"(-?\d+)(\.\d+)?"; // 匹配一串连续的数字  

            Regex regex = new Regex(pattern);
            System.Text.RegularExpressions.Match match = regex.Match(numberStr);

            if (match.Success)
            {
                return int.Parse(match.Value);
            }
            else
            {
                return 0;
            }
        }
    }
}
