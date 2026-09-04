using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace out_exel
{
    
    internal class DBHelper
    {
        //static string Conn = @"Data Source=192.168.0.213;Initial Catalog=iMES;User ID=sa;Password=SXsd123!@#";
        //public static string Conn = System.Configuration.ConfigurationManager.AppSettings["connectstrsql"];
        static string Conn = "data source=localhost;database=test;user id=root;password=123456;pooling=true;charset=utf8;";

        public static DataTable ExecuterQuery(string commandSql)
        {
            return ExecuterQuery(Conn, commandSql);
        }

        public static DataTable ExecuterQuery(string connectionString, string commandSql)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection sqlConnection =
                new SqlConnection(connectionString))
                {
                    sqlConnection.Open();

                    using (SqlDataAdapter sqlDataAdapter =
                    new SqlDataAdapter(commandSql, sqlConnection))
                    {
                        sqlDataAdapter.Fill(dataTable);
                    }

                    sqlConnection.Close();
                }
            }
            catch (Exception ex)
            {
                return null;
            }

            return dataTable;
        }
        public static DataTable ExecuterQuery(string connectionString, string commandSql, ref string error)
        {
            DataTable dataTable = new DataTable();

            try
            {
                using (SqlConnection sqlConnection =
                new SqlConnection(connectionString))
                {
                    sqlConnection.Open();

                    using (SqlDataAdapter sqlDataAdapter =
                    new SqlDataAdapter(commandSql, sqlConnection))
                    {
                        sqlDataAdapter.Fill(dataTable);
                    }

                    sqlConnection.Close();
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }

            return dataTable;
        }

        public static int ExecuterNoQuery(string commandSql)
        {
            return ExecuterNoQuery(Conn, commandSql);
        }


        public static int ExecuterNoQuery(string connectionString, string commandSql)
        {
            int affectRow = 0;
            SqlConnection sqlConnection = null;
            try
            {
                using (sqlConnection = new SqlConnection(connectionString))
                {
                    sqlConnection.Open();

                    using (SqlCommand sqlCommand = new SqlCommand(commandSql, sqlConnection))
                    {
                        affectRow = sqlCommand.ExecuteNonQuery();
                    }

                    sqlConnection.Close();
                }
            }
            catch (Exception ex)
            {
                if (sqlConnection != null && sqlConnection.State == ConnectionState.Open)
                {
                    sqlConnection.Close();
                }
                if (ex.Message.Contains("并且已被选作死锁牺牲品"))
                {
                    affectRow = ExecuterNoQuery(commandSql);
                }
                else
                {
                    return 0;
                }
            }

            return affectRow;
        }

        public static int ExecuterNoQuery(string connectionString, string commandSql, ref string error)
        {
            int affectRow = 0;
            SqlConnection sqlConnection = null;
            try
            {
                using (sqlConnection = new SqlConnection(connectionString))
                {
                    sqlConnection.Open();

                    using (SqlCommand sqlCommand = new SqlCommand(commandSql, sqlConnection))
                    {
                        affectRow = sqlCommand.ExecuteNonQuery();
                    }

                    sqlConnection.Close();
                }
            }
            catch (Exception ex)
            {
                if (sqlConnection != null && sqlConnection.State == ConnectionState.Open)
                {
                    sqlConnection.Close();
                }
                if (ex.Message.Contains("并且已被选作死锁牺牲品"))
                {
                    affectRow = ExecuterNoQuery(commandSql);
                }
                else
                {
                    error = ex.Message;
                    return 0;
                }
            }

            return affectRow;
        }

        public static bool ExecuterNoQueryByTran(SqlConnection sqlConnection, SqlTransaction sqlTransaction, string commandSql, out string exMSG)
        {
            exMSG = "";
            int affectRow = 0;
            try
            {

                using (SqlCommand sqlCommand = new SqlCommand(commandSql, sqlConnection))
                {
                    sqlCommand.Connection = sqlConnection;
                    sqlCommand.Transaction = sqlTransaction;
                    affectRow = sqlCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                exMSG = ex.Message;
                return false;
            }

            return true;
        }




        /// <summary>
        /// DataTableToJson
        /// </summary>
        /// <param name="jsonName"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string DataTableToJson(string jsonName, DataTable dt)
        {
            StringBuilder Json = new StringBuilder();
            Json.Append("{\"" + jsonName + "\":[");
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Json.Append("{");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        Json.Append("\"" + dt.Columns[j].ColumnName.ToString() + "\":\"" + dt.Rows[i][j].ToString() + "\"");
                        if (j < dt.Columns.Count - 1)
                        {
                            Json.Append(",");
                        }
                    }
                    Json.Append("}");
                    if (i < dt.Rows.Count - 1)
                    {
                        Json.Append(",");
                    }
                }
            }
            Json.Append("]}");
            return Json.ToString();
        }
    }


}
