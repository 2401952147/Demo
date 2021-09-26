using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;

namespace Demo.Common
{
    public class DBHelper
    {
        private static readonly string ConnStr = "Data Source=.;Initial Catalog=Test;Integrated Security=True";
        /// <summary>
        /// 返回datatable类型，一般用于查询指定条件的数据
        /// </summary>
        /// <returns></returns>
        public DataSet ExecuteQuery(string sql) {
            using (SqlConnection conn = new SqlConnection(ConnStr))
            {
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SqlDataAdapter ds = new SqlDataAdapter(sql, conn);
                DataSet dataSet = new DataSet();
                ds.Fill(dataSet);
                conn.Close();
                return dataSet;
            }
        }

        /// <summary>
        /// 返回bool类型，一般用于增，删，改等操作
        /// </summary>
        /// <param name="sql"></param>
        /// <returns></returns>
        public bool ExecuteNonQuery(string sql) {
            using (SqlConnection conn = new SqlConnection(ConnStr))
            {
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SqlCommand cmd = new SqlCommand(sql, conn);
                if (cmd.ExecuteNonQuery() > 0)
                {
                    conn.Close();
                    return true;
                }
                else
                {
                    conn.Close();
                    return false;
                }
            }
        }
    }
}