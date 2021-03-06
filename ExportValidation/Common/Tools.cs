﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;

using System.IO;
using System.Xml;

namespace ExportValidation.Common
{
    static class Tools
    {
        private static string columnNames(DataTable dtSchemaTable, string delimiter)
        {
            string strOut = "";
            if (delimiter.ToLower() == "tab")
            {
                delimiter = "\t";
            }

            for (int i = 0; i < dtSchemaTable.Rows.Count; i++)
            {
                strOut += dtSchemaTable.Rows[i][0].ToString();
                if (i < dtSchemaTable.Rows.Count - 1)
                {
                    strOut += delimiter;
                }

            }
            return strOut;
        }

        private static List<string> GetNamesFromSQL(SqlConnection conn, string sql)
        {

            var lst = new List<string>();
            var cmd = new SqlCommand(sql, conn);

            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }

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
            var sql = "SELECT LOWER(name) AS name FROM sysobjects WHERE xtype='P' AND name NOT LIKE 'dt_%' ORDER BY NAME";
            var lst = GetNamesFromSQL(conn, sql);

            return lst;
        }

        public static List<IndexData> GetIndex(SqlConnection conn, string procname)
        {
            var sql = "exec " + procname;
            var cmd = new SqlCommand(sql, conn);
            var lst = new List<IndexData>();
            var lstData = new List<IndexData>();

            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }
            var rdr = cmd.ExecuteReader();

            // Получаем список запросов валидации
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    lst.Add(new IndexData
                    {
                        ValidationRule = rdr.GetString(0),
                        NameList = rdr.GetString(3),
                        Description = rdr.GetString(2),
                        SelectCommand = rdr.GetString(5)
                    });
                }
            }
            rdr.Close();
            rdr = null;


            return lst;
        }

        public static List<IndexData> GetIndexWithSelectCommand(SqlConnection conn, string procname)
        {
            var sql = "exec " + procname;
            var cmd = new SqlCommand(sql, conn);
            var lst = new List<IndexData>();
            var lstData = new List<IndexData>();

            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }
            var rdr = cmd.ExecuteReader();

            // Получаем список запросов валидации
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    lst.Add(new IndexData
                    {
                        ValidationRule = rdr.GetString(1),
                        NameList = rdr.GetString(4),
                        Description = rdr.GetString(3),
                        SelectCommand = rdr.GetString(5)
                    });
                }
            }
            rdr.Close();
            rdr = null;


            return lst;
        }

        private static QueryData GetQueryData(string procName, string validName, string selectText, string descText, string number, string execText, string projectName, SqlConnection conn)
        {
            var obj = new QueryData();
            var lstColumns = new List<string>();
            var sql = execText;
            var cmd = new SqlCommand(sql.Replace("\\r\\n", "").Replace("\\r", ""), conn);
            if (cmd.CommandText.StartsWith("EXEC"))
            {
                cmd.CommandTimeout = 200;
            }
            if (cmd.CommandText.Contains("check_patient"))
            {
                cmd.CommandTimeout = 90;
            }
            cmd.CommandTimeout = 200;
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

        public static void RunProcedureNonQuery(SqlConnection conn, string procedureName)
        {
            MessageBox.Show("Start: " + procedureName);
            var sql = "EXEC " + procedureName;
            var cmd = new SqlCommand(sql, conn);
            try
            {
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }

                cmd.ExecuteNonQuery();


            }
            catch (Exception e)
            {

                MessageBox.Show(e.Message + " - " + procedureName + " - ");
            }
            finally
            {
                MessageBox.Show("Exit: " + procedureName);
            }

        }


        public static List<QueryData> RunProcedure(SqlConnection conn, string procedureName, string projectName, DateTime? startdate = null, DateTime? enddate = null)
        {
            MessageBox.Show("Start: " + procedureName);
            var sql = "EXEC " + procedureName;
            var cmd = new SqlCommand(sql, conn);
            var lst = new List<ValidationRows>();
            var lstData = new List<QueryData>();
            try
            {
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }

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

                if (startdate != null && enddate != null)
                {
                    foreach (var validationRowse in lst)
                    {
                        validationRowse.s6 = validationRowse.s6 +
                        " @start_date = N'" + startdate.ToString() + "', @end_date = N'" + enddate.ToString() + "'";
                    }
                }
                else
                {
                    foreach (var validationRowse in lst)
                    {
                        if (validationRowse.s6.Contains("GET_USER"))
                        {
                            validationRowse.s6 = validationRowse.s6 +
                                                 " @start_date = NULL, @end_date = NULL";
                        }
                    }
                }


                foreach (var item in lst)
                {
                    var obj = GetQueryData(item.s2, item.s1, item.s5, item.s3, item.s4, item.s6, projectName, conn);
                    if (obj != null)
                    {
                        lstData.Add(obj);
                    }
                }
            }
            catch (Exception e)
            {


                MessageBox.Show(e.Message + " - " + procedureName + " - ");
            }
            finally
            {
                MessageBox.Show("Exit: " + procedureName);
            }

            return lstData;
        }

        public static void GetQueries(SqlConnection conn, string strProject, string strPath)
        {
            var sql = "EXEC	[dbo].[GetQueries]";
            var cmd = new SqlCommand(sql, conn);
            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }
            cmd.CommandTimeout = 120;
            cmd.ExecuteNonQuery();

            sql =
                "SELECT DISTINCT [UserName],[UserEmail],[SiteNo],[CityName] FROM [dbo].[QUERY_LIST_DISTINCT]";

            var lstUsers = new List<QueryReportData>();
            cmd = new SqlCommand(sql, conn);
            var rdr = cmd.ExecuteReader();
            // Получаем список пользователей
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    lstUsers.Add(new QueryReportData
                    {
                        UserName = rdr.GetString(0),
                        UserEmail = rdr.GetString(1),
                        SiteNo = rdr.GetString(2),
                        CityName = rdr.GetString(3)
                    });
                }
            }
            rdr.Close();
            rdr = null;



            foreach (var user in lstUsers)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(strPath + @"\" + strProject + "_" + user.SiteNo + "_" + user.CityName + "_" + user.UserName + DateTime.Now.ToShortDateString() + ".txt", true))
                {
                    file.WriteLine("Номер центра: " + user.SiteNo);
                    file.WriteLine("Город: " + user.CityName);
                    file.WriteLine("Пользователь: " + user.UserName);
                    file.WriteLine("Email: " + user.UserEmail);
                    file.WriteLine("");
                    file.WriteLine("");
                    file.WriteLine("");

                    sql = "SELECT DISTINCT [CrfNumber],[CrfName],[DateOfInput]  FROM [dbo].[QUERY_LIST_DISTINCT]  WHERE UserName = N'" + user.UserName + "'";
                    cmd = new SqlCommand(sql, conn);
                    rdr = cmd.ExecuteReader();
                    var lstCrf = new List<CrfInfo>();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            lstCrf.Add(new CrfInfo
                            {
                                CrfNumber = rdr.GetString(0),
                                CrfName = rdr.GetString(1),
                                DateOfInput = rdr.GetDateTime(2).ToString()
                            });
                        }
                    }
                    rdr.Close();
                    rdr = null;
                    foreach (var crfInfo in lstCrf)
                    {
                        file.WriteLine("");
                        file.WriteLine("----------------------------");
                        file.WriteLine("");
                        file.WriteLine("Номер карты: " + crfInfo.CrfNumber);
                        file.WriteLine("Пациент: " + crfInfo.CrfName);
                        file.WriteLine("Дата ввода: " + crfInfo.DateOfInput);
                        file.WriteLine("");
                        file.WriteLine("Правила валидации и описание ошибки:");

                        sql =
                            "SELECT DISTINCT [ValidationRule],[Descritpion]  FROM [dbo].[QUERY_LIST_DISTINCT] WHERE UserName = N'" +
                            user.UserName + "' AND CrfNumber = N'" + crfInfo.CrfNumber + "'";
                        cmd = new SqlCommand(sql, conn);
                        rdr = cmd.ExecuteReader();
                        if (rdr.HasRows)
                        {
                            while (rdr.Read())
                            {
                                {
                                    file.WriteLine(rdr.GetString(0) + " - " + rdr.GetString(1));
                                }
                            }
                        }
                        rdr.Close();
                        rdr = null;
                    }
                    file.WriteLine("");
                    file.WriteLine("----------------------------");
                }

            }
            MessageBox.Show("Создание Квери закончено");


        }
        public static void GetQueriesInFormat(SqlConnection conn, string strProject, string strPath)
        {
            var sql = "EXEC	[dbo].[GetQueries]";
            var cmd = new SqlCommand(sql, conn);
            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }
            try
            {
                cmd.CommandTimeout = 120;
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {

            }
            sql =
                "SELECT DISTINCT [UserName],[UserEmail],[SiteNo],[CityName] FROM [dbo].[QUERY_LIST_DISTINCT]";

            var lstUsers = new List<QueryReportData>();
            cmd = new SqlCommand(sql, conn);
            var rdr = cmd.ExecuteReader();
            // Получаем список пользователей
            if (rdr.HasRows)
            {
                while (rdr.Read())
                {
                    lstUsers.Add(new QueryReportData
                    {
                        UserName = rdr.GetString(0),
                        UserEmail = rdr.GetString(1),
                        SiteNo = rdr.GetString(2),
                        CityName = rdr.GetString(3)
                    });
                }
            }
            rdr.Close();
            rdr = null;



            foreach (var user in lstUsers)
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(strPath + @"\" + strProject + "_" + user.SiteNo + "_" + user.CityName + "_" + user.UserName + "_" + DateTime.Now.ToShortDateString() + ".txt", true))
                {
                    file.WriteLine("To: " + user.UserEmail);
                    file.WriteLine("Subject: " + strProject + " Query");
                    file.WriteLine("");
                    file.WriteLine("");
                    file.WriteLine("Уважаемый(-ая), " + user.UserName + ",");

                    sql = "SELECT DISTINCT [CrfNumber],[CrfName],[DateOfInput]  FROM [dbo].[QUERY_LIST_DISTINCT]  WHERE UserName = N'" + user.UserName + "'";
                    cmd = new SqlCommand(sql, conn);
                    rdr = cmd.ExecuteReader();
                    var lstCrf = new List<CrfInfo>();
                    if (rdr.HasRows)
                    {
                        while (rdr.Read())
                        {
                            lstCrf.Add(new CrfInfo
                            {
                                CrfNumber = rdr.GetString(0),
                                CrfName = rdr.GetString(1)
                            });
                        }
                    }
                    rdr.Close();
                    rdr = null;
                    foreach (var crfInfo in lstCrf)
                    {
                        file.WriteLine("");

                        file.WriteLine("В карте № " + crfInfo.CrfNumber + " пациента " + crfInfo.CrfName + " обнаружены проблемные данные:\r\n");

                        sql =
                            "SELECT DISTINCT [ValidationRule],[Descritpion]  FROM [dbo].[QUERY_LIST_DISTINCT] WHERE UserName = N'" +
                            user.UserName + "' AND CrfNumber = N'" + crfInfo.CrfNumber + "'";
                        cmd = new SqlCommand(sql, conn);
                        rdr = cmd.ExecuteReader();
                        var i = 1;
                        if (rdr.HasRows)
                        {
                            while (rdr.Read())
                            {
                                {
                                    file.WriteLine(i + ") " + rdr.GetString(1));
                                    i++;
                                }
                            }
                        }
                        rdr.Close();
                        rdr = null;
                        file.WriteLine("");
                    }
                    file.WriteLine("\r\n\r\n\r\nПожалуйста, в случае ошибки, исправьте их самостоятельно или свяжитесь со службой поддержки путем ответа на данное письмо. \r\nПожалуйста, не удаляйте тему письма при ответе!\r\n" +
                                   "Крайний срок внесения исправлений 8:00 14.08.2014 (четверг)\r\n\r\nСпасибо,\r\nКоманда поддержки eCRF");
                }

            }
            MessageBox.Show("Создание Квери закончено");


        }

        public static List<string> GetTablesInDataBase(SqlConnection conn)
        {
            var lst = new List<string>();

            SqlDataAdapter da = new SqlDataAdapter("select name from dbo.sysobjects where xtype = 'U' and name <> 'dtproperties' order by name", conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("There is no user table in the specified database.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lst.Add("<no tables>");
            }
            else
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    lst.Add(dt.Rows[i][0].ToString());
                }
            }

            return lst;
        }

        public static List<string> GetViewsInDataBase(SqlConnection conn)
        {
            var lst = new List<string>();

            SqlDataAdapter da = new SqlDataAdapter("select  name from    sys.objects where type = 'V' and name <> 'dtproperties' order by name", conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("There is no user views in the specified database.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                lst.Add("<no tables>");
            }
            else
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    lst.Add(dt.Rows[i][0].ToString());
                }
            }

            return lst;


        }

        public static void ExportToXMLFile(string filePath, string project, string filename, string sql,
            SqlConnection conn)
        {
            var cmd = sql + " FOR XML PATH('" + filename + "'), ROOT ('" + filename + "s')";

         //   cmd.CommandTimeout = 240;
            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }
            var adapter = new SqlDataAdapter(sql, conn);
            var ds = new DataSet();
            adapter.Fill(ds);
            ds.WriteXml(filePath + "\\" + filename + ".xml");
            conn.Close();
        }

        public static void ExportToCSVFile(string filePath, string project, string fileName, string sql, SqlConnection conn, Encoding encoding, string separator, bool firstRowNames)
        {
            var cmd = new SqlCommand(sql, conn);
            cmd.CommandTimeout = 240;
            if (conn.State.ToString() == "Closed")
            {
                conn.Open();
            }
            // Creates a SqlDataReader instance to read data from the table.
            SqlDataReader dr = cmd.ExecuteReader();

            // Retrives the schema of the table.
            DataTable dtSchema = dr.GetSchemaTable();

            // Creates the CSV file as a stream, using the given encoding.
            //StreamWriter sw = new StreamWriter(filePath + "\\" + project + "_" + fileName + "_" + DateTime.Now.ToShortDateString() + ".csv", false, encoding);
            StreamWriter sw = new StreamWriter(filePath + "\\" + fileName  + ".csv", false, encoding);
            StringBuilder strRow; // represents a full row

            // Writes the column headers if the user previously asked that.
            if (firstRowNames)
            {
                sw.WriteLine(columnNames(dtSchema, separator));
            }

            // Reads the rows one by one from the SqlDataReader
            // transfers them to a string with the given separator character and
            // writes it to the file.
            while (dr.Read())
            {
                strRow = new StringBuilder();
                for (int i = 0; i < dr.FieldCount; i++)
                {
                    strRow.Append(dr.GetValue(i).ToString());
                    if (i < dr.FieldCount - 1)
                    {
                        strRow.Append(separator);
                    }
                }
                sw.WriteLine(strRow);
            }


            // Closes the text stream and the database connenction.
            sw.Close();
            conn.Close();
        }

        public static void GenerateCSVDocument(SqlConnection conn, List<IndexData> index, string project, string filePath, Encoding encoding, string separator, bool hasColumnNames)
        {
            foreach (var indexData in index)
            {
                ExportToCSVFile(filePath, project, indexData.NameList, indexData.SelectCommand, conn, encoding, separator, hasColumnNames);
            }
        }

        public static void ExportToDBFFile(string filePath, string strProject, string filename, string sqldata, SqlConnection conn)
        {
         

            try
            {
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }
                // create linked server
                var sql = @"EXEC master.dbo.sp_addlinkedserver @server = N'VFP_Test', @srvproduct=N'Microsoft Visual FoxPro OLE DB Data Provider', @provider=N'VFPOLEDB', @datasrc=N'"+filePath+"',@provstr=N'VFPOLEDB.1'";
                var cmd = new SqlCommand(sql, conn);
                cmd.ExecuteNonQuery();

                sql = "INSERT INTO VFP_Test." + filename + " " + sqldata;
                cmd = new SqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                // transfer data

                // delete linked server
                sql = @"IF  EXISTS (SELECT srv.name FROM sys.servers srv WHERE srv.server_id != 0 AND srv.name = N'VFP_Test')EXEC master.dbo.sp_dropserver @server=N'VFP_Test', @droplogins='droplogins'";
                cmd = new SqlCommand(sql, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
      }
    }
}
