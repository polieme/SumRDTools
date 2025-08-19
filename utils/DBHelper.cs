using System.Data;
using System.Data.SQLite;

namespace SumRDTools
{


    public static class DBHelper
    {
        // 数据库连接字符串（自动创建数据库文件）
        private static string ConnectionString = "Data Source=./sumrdtools.db;Version=3;";

        /// <summary>
        /// 获取数据库连接
        /// </summary>
        private static SQLiteConnection GetConnection()
        {
            return new SQLiteConnection(ConnectionString);
        }

        /// <summary>
        /// 执行非查询操作（增删改）
        /// </summary>
        public static int ExecuteNonQuery(string sql, params SQLiteParameter[] parameters)
        {
            using (var conn = GetConnection())
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddRange(parameters);
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 执行查询，返回DataTable
        /// </summary>
        public static DataTable ExecuteQuery(string sql, params SQLiteParameter[] parameters)
        {
            using (var conn = GetConnection())
            {
                conn.Open();
                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddRange(parameters);
                    using (SQLiteDataReader reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        return dt;
                    }
                }
            }
        }

        // 可选：其他封装方法（如ExecuteScalar等）
    }
}
