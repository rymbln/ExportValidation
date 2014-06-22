using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportValidation.Common
{
    static class Tools

        //            using (System.IO.StreamWriter file = new System.IO.StreamWriter(filePath))
    // {
    {


        private static List<string> GetNamesFromSQL(SqlConnection conn, string sql)
        {
            var lst = new List<string>();
            var cmd = new SqlCommand(sql, conn);
            conn.Open();

            var rdr = cmd.ExecuteReader();

            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    lst.Add(rdr.GetString(0));
                }
            }
            return lst;
        }

        public static SqlConnection GetConnectionString(string serverName, string loginName, string passwordName)
        {
            string connStr = "Server=" + serverName + "; User Id=" + loginName + "; password= " + passwordName;
            var conn = new SqlConnection(connStr);

            return conn;
        }

        public static SqlConnection GetConnectionString(string serverName, string dbName, string loginName, string passwordName)
        {
            string connStr = "Server=" + serverName + "; Initial Catalog=" + dbName + "; User Id=" + loginName + "; password= " + passwordName;
            var conn = new SqlConnection(connStr);

            return conn;
        }


        public static List<string> GetDatabaseNames(SqlConnection conn)
        {
            var sql = "SELECT name FROM sys.databases";
            var lst = GetNamesFromSQL(conn, sql);

            return lst;
        }

        public static List<string> GetProceduresInDatabase(SqlConnection conn)
        {
            var sql = "SELECT LOWER(name) AS name FROM sysobjects WHERE xtype='P' AND name NOT LIKE 'dt_%'";
            var lst = GetNamesFromSQL(conn, sql);

            return lst;
        }

        private static QueryData GetQueryData(string procName, string validName, string selectText, string descText, string number, string execText, string projectName, SqlConnection conn)
        {
            var obj = new QueryData();
            var lstColumns = new List<string>();
            var sql = execText;
            var cmd = new SqlCommand(sql, conn);
            var rdrQD = cmd.ExecuteReader();
            if (rdrQD.HasRows)
            {
                for (int i = 0; i < rdrQD.FieldCount; i++)
                {
                    lstColumns.Add(rdrQD.GetName(i));
                }

                obj.FieldsName = lstColumns;
                // Получаем сами данные
                var dtData = new DataTable();
                dtData.Load(rdrQD);

                obj.Data = dtData;
                obj.ProjectName = projectName;
                obj.ValidationRule = validName;
                obj.Description = descText;
                obj.NameList = number;
            }
            else
            {
                obj = null;
            }

            rdrQD.Close();
            rdrQD = null;

            return obj;
        }

        public static List<QueryData> RunProcedure(SqlConnection conn, string procedureName, string projectName)
        {
            var sql = "EXEC " + procedureName;
            var cmd = new SqlCommand(sql, conn);
            var lst = new List<ValidationRows>();
            var lstData = new List<QueryData>();
            FileStream fs = new FileStream(@"C:\Users\rymbln\Desktop\WriteLines2.txt", FileMode.OpenOrCreate, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);

            conn.Open();

            var rdr = cmd.ExecuteReader();

            // Получаем список запросов валидации
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    lst.Add(new ValidationRows
                    {

                        s1 = rdr.GetString(0),
                        s2 = rdr.GetString(1),
                        s3 = rdr.GetString(2),
                        s4 = rdr.GetString(3),
                        s5 = rdr.GetString(4),
                        s6 = rdr.GetString(5),
                    });
                }
            }
            rdr.Close();
            rdr = null;

            foreach (var item in lst)
            {
                var obj = GetQueryData(item.s1, item.s2, item.s3, item.s4, item.s5, item.s6, projectName, conn);
                
                if (obj != null)
                {
                lstData.Add(obj);
                }
            }
            return lstData;
        }
    }
}
