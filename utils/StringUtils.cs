using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SumRDTools
{
    public class StringUtils
    {


        //正则表达式获取字符串中的英文字母
        public static String getContainsChar(string charStr)
        {

            //通过正则获取到表格中的数据
            string pattern = @"[A-Za-z]{2,100}"; // 匹配字符串中包含一连串英文的项目  

            Regex regex = new Regex(pattern);
            System.Text.RegularExpressions.Match match = regex.Match(charStr);

            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return "";
            }
        }
    }
}
