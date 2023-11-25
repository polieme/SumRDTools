using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SumRDTools
{
    public static class FileOptUtils
    {

        //删除指定文件夹下的文件
        public static void DeleteAllFiles(string folderPath)
        {
            DirectoryInfo directory = new DirectoryInfo(folderPath);
            FileInfo[] files = directory.GetFiles();

            foreach (FileInfo file in files)
            {
                if (file.Exists)
                {
                    file.Delete();
                }
            }
        }
    }
}
