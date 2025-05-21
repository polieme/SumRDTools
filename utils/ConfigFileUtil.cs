using NPOI.OpenXmlFormats.Vml;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SumRDTools
{
    public static class ConfigFileUtil
    {
        //配置文件默认路径
        private static string configFilePath = "./SumRDTools.ini";

        [DllImport("kernel32")]// 读配置文件方法的6个参数：所在的分区（section）、 键值、     初始缺省值、   StringBuilder、  参数长度上限 、配置文件路径
        public static extern long GetPrivateProfileString(string section, string key, string defaultValue, StringBuilder retVal, int size, string filePath);
        [DllImport("kernel32")]//写入配置文件方法的4个参数：  所在的分区（section）、  键值、     参数值、       配置文件路径
        private static extern long WritePrivateProfileString(string section, string key, string value, string filePath);


        //读取配置文件参数
        public static string getConfigParam(String section, String key)
        {
            //获取程序当前所在目录
            if (File.Exists(configFilePath))  //检查是否有配置文件，并且配置文件内是否有相关数据。
            {
                StringBuilder paramsStringBulder= new StringBuilder(255);
                GetPrivateProfileString(section, key, "配置文件不存在，读取未成功!", paramsStringBulder, 255, configFilePath);
                return paramsStringBulder.ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        /*写配置文件*/
        public static void SetValue(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, configFilePath);
        }

    }
}
