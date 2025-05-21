using System;

namespace SumRDTools
{
    public class DateUtils
    {
        private static String Format_yyyyMM = "yyyyMM";
        private static String Format_yyyyMMdd = "yyyyMMdd";


        //格式化日期
        public static DateTime formatDatetime(String datetimeStr) {
            DateTime result;
            if (DateTime.TryParseExact(datetimeStr, Format_yyyyMM, null, System.Globalization.DateTimeStyles.None, out result))
            {
                return result;
            }
            if (DateTime.TryParseExact(datetimeStr, Format_yyyyMMdd, null, System.Globalization.DateTimeStyles.None, out result))
            {
                return result;
            }
            else if (double.TryParse(datetimeStr,out Double date))
            {
                return DateTime.FromOADate(date);
            }
            else {
                Console.WriteLine("解析失败："+datetimeStr);
            }

            return new DateTime();
        }
    }
}
