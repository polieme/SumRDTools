using System.Text.RegularExpressions;

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

            decimal decimalVal = getDecimal(numberStr);

            return int.Parse(decimalVal.ToString("0"));
        }
    }
}
