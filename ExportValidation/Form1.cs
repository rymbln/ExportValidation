using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExportValidation.Common;

namespace ExportValidation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string strFormat;	//CSV separator character
        private string strEncoding; //Encoding of CSV file
        private string fileCSV;		//full file name
        private string dirCSV;		//directory of file to import
        private string fileNevCSV;	//name (with extension) of file to import - field

        public string FileNevCSV	//name (with extension) of file to import - property
        {
            get { return fileNevCSV; }
            set { fileNevCSV = value; }
        }
        private string separatorCSV
        {
            get
            {
                if (rdbSemicolon.Checked)
                {
                    return ";";
                }
                else if (rdbTab.Checked)
                {
                    return "\t";
                }
                else if (rdbSeparatorOther.Checked)
                {
                    if (txtSeparatorOtherChar.Text.Length == 1)
                    {
                        return txtSeparatorOtherChar.Text;
                    }
                    else
                    {
                        MessageBox.Show("Invalid separator character.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        rdbSemicolon.Checked = true;
                        return ";";
                    }
                }
                else
                {
                    return ";";
                }
            }
        }


        private void ImportEncoding()
        {
            try
            {

                if (rdbImportAnsi.Checked)
                {
                    strEncoding = "ANSI";
                }
                else if (rdbImportUnicode.Checked)
                {
                    strEncoding = "Unicode";
                }
                else if (rdbImportOEM.Checked)
                {
                    strEncoding = "OEM";
                }
                else
                {
                    strEncoding = "ANSI";
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Encoding");
            }
            finally
            {
            }
        }
        private Encoding encodingCSV
        {
            get
            {
                if (rdbUnicode.Checked)
                {
                    return Encoding.Unicode;
                }
                else if (rdbASCII.Checked)
                {
                    return Encoding.ASCII;
                }
                else if (rdbUTF7.Checked)
                {
                    return Encoding.UTF7;
                }
                else if (rdbUTF8.Checked)
                {
                    return Encoding.UTF8;
                }
         else
                {
                    return Encoding.Unicode;
                    
                }

                // You can add other options, for ex.:
                //return Encoding.GetEncoding("iso-8859-2");
                //return Encoding.Default;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;

            this.cbxDatabases.Items.Clear();

            var conn = Tools.GetConnectionString(strServer, strLogin, strPassword);
            using (conn)
            {
                var lst = Tools.GetDatabaseNames(conn);
                foreach (var item in lst)
                {
                    this.cbxDatabases.Items.Add(item);
                }
                this.cbxDatabases.Refresh();
                this.statusStrip1.Text = "Список баз данных получен";
                this.statusStrip1.Refresh();

            }
        }

        private void btnProcedures_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();

            this.cbxProcedures.Items.Clear();

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);
            using (conn)
            {
                var lst = Tools.GetProceduresInDatabase(conn);
                foreach (var item in lst)
                {
                    this.cbxProcedures.Items.Add(item);
                }
                this.cbxProcedures.Refresh();
                this.statusStrip1.Text = "Список процедур получен";
                this.statusStrip1.Refresh();

            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);

            using (conn)
            {
                var data = Tools.RunProcedure(conn, this.cbxProcedures.SelectedItem.ToString(), strProject);
                if (data.Count > 0)
                {
                    PDFGeneration.GenerateDocument(strPath, data);
                    MessageBox.Show("Finish");
                }
                else
                {
                    MessageBox.Show("Ошибок не обнаружено!");
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);

            using (conn)
            {
                var data = Tools.RunProcedure(conn, "_ExportDataValidation", strProject);
                var index = Tools.GetIndex(conn, "_ExportDataValidation");
                if (data.Count > 0)
                {
                    ExcelGeneration.GenerateDocument(strPath, data, index);
                    MessageBox.Show("Finish");
                }
                else
                {
                    MessageBox.Show("Ошибок не обнаружено!");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;
            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);

            using (conn)
            {
                var data = Tools.RunProcedure(conn, this.cbxProcedures.SelectedItem.ToString(), strProject);
                if (data.Count > 0)
                {
                    WordGeneration.GenerateDocument(strPath, data, "portrait");
                    MessageBox.Show("Finish");
                }
                else
                {
                    MessageBox.Show("Ошибок не обнаружено!");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;
            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);

            using (conn)
            {
                var data = Tools.RunProcedure(conn, this.cbxProcedures.SelectedItem.ToString(), strProject);
                if (data.Count > 0)
                {
                    WordGeneration.GenerateDocument(strPath, data, "album");
                    MessageBox.Show("Finish");
                }
                else
                {
                    MessageBox.Show("Ошибок не обнаружено!");
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);

            using (conn)
            {
                var data = Tools.RunProcedure(conn, "_ExportDataForStatistic", strProject);
                var index = Tools.GetIndex(conn, "_ExportDataForStatistic");
                if (data.Count > 0)
                {
                    ExcelGeneration.GenerateDocument2(strPath, data, index);
                    MessageBox.Show("Finish");
                }
                else
                {
                    MessageBox.Show("Ошибок не обнаружено!");
                }
            }
        }

        private void cbxDatabases_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.tbxProjectName.Text = this.cbxDatabases.SelectedItem.ToString();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);

            using (conn)
            {
                Tools.GetQueries(conn, strProject, strPath);

            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);

            using (conn)
            {
                Tools.GetQueriesInFormat(conn, strProject, strPath);

            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            folderBrowserDialog.SelectedPath = @"C:\Users\rymbln\Desktop\Output\";
            //  openFileDialogCSV.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            //  openFileDialogCSV.FilterIndex = 1;

            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                this.tbxOutputPath.Text = folderBrowserDialog.SelectedPath.ToString();
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();

            this.lbxTables.Items.Clear();
            this.lbxViews.Items.Clear();
            this.lbxProcedures.Items.Clear();
            
            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);
            using (conn)
            {
                var lst = Tools.GetTablesInDataBase(conn);
                foreach (var item in lst)
                {
                    this.lbxTables.Items.Add(item);
                }
                this.lbxTables.Refresh();

                var lst2 = Tools.GetViewsInDataBase(conn);
                foreach (var item in lst2)
                {
                    this.lbxViews.Items.Add(item);
                }
                this.lbxViews.Refresh();

                var lst3 = Tools.GetProceduresInDatabase(conn);
                foreach (var item in lst3)
                {
                    this.lbxProcedures.Items.Add(item);
                }
                this.lbxProcedures.Refresh();
            }
        }

        private void lbl01_Click(object sender, EventArgs e)
        {

        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.lbxProcedures.SelectedItems.Clear();
            this.lbxViews.SelectedItems.Clear();
            this.lbxTables.SelectedItems.Clear();
        }

        private void btnExportToCSV_Click(object sender, EventArgs e)
        {
            string fileName = "";
            string sql = "";

              var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();
            var strPath = this.tbxOutputPath.Text;
            var strProject = this.tbxProjectName.Text;

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);
            if (lbxProcedures.SelectedItems.Count > 0)
            {
                foreach (var selectedItem in lbxProcedures.SelectedItems)
                {
                    fileName = selectedItem.ToString();
                    sql = "EXEC " + fileName;
                    Tools.ExportToCSVFile(strPath, strProject, fileName, sql, conn, encodingCSV, separatorCSV, this.chkFirstRowColumnNames.Checked);
                }
            }
            if (lbxViews.SelectedItems.Count > 0)
            {
                foreach (var selectedItem in lbxViews.SelectedItems)
                {
                    fileName = selectedItem.ToString();
                    sql = "SELECT * FROM " + fileName;
                    Tools.ExportToCSVFile(strPath, strProject, fileName, sql, conn, encodingCSV, separatorCSV, this.chkFirstRowColumnNames.Checked);
                }
            }
            if (lbxTables.SelectedItems.Count > 0)
            {
                foreach (var selectedItem in lbxTables.SelectedItems)
                {
                    fileName = selectedItem.ToString();
                    sql = "SELECT * FROM " + fileName;
                    Tools.ExportToCSVFile(strPath, strProject, fileName, sql, conn, encodingCSV, separatorCSV, this.chkFirstRowColumnNames.Checked);
                }
            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            this.lbxTables.SelectedItems.Add(lbxTables.Items);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            this.lbxViews.SelectedItems.Add(lbxViews.Items);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            this.lbxProcedures.SelectedItems.Add(lbxProcedures.Items);
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogCSV = new OpenFileDialog();

            openFileDialogCSV.InitialDirectory = Application.ExecutablePath.ToString();
            openFileDialogCSV.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialogCSV.FilterIndex = 1;
            openFileDialogCSV.RestoreDirectory = true;

            if (openFileDialogCSV.ShowDialog() == DialogResult.OK)
            {
                this.txtFileToImport.Text = openFileDialogCSV.FileName.ToString();
            }
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            loadPreview();
        }

        private void loadPreview()
        {
            try
            {
                // select format, encoding, an write the schema file
                Format();
                ImportEncoding();
                writeSchema();

                // loads the first 500 rows from CSV file, and fills the
                // DataGridView control.
                this.dataGridView_preView.DataSource = LoadCSV(500);
                this.dataGridView_preView.DataMember = "csv";
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error - loadPreview", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void writeSchema()
        {
            try
            {
                FileStream fsOutput = new FileStream("D:\\schema.ini", FileMode.Create, FileAccess.Write);
                StreamWriter srOutput = new StreamWriter(fsOutput);
                string s1, s2, s3, s4, s5;

                s1 = "[" + this.FileNevCSV + "]";
                s2 = "ColNameHeader=" + chkFirstRowColumnNames.Checked.ToString();
                s3 = "Format=" + this.strFormat;
                s4 = "MaxScanRows=25";
                s5 = "CharacterSet=" + this.strEncoding;

                srOutput.WriteLine(s1.ToString() + "\r\n" + s2.ToString() + "\r\n" + s3.ToString() + "\r\n" + s4.ToString() + "\r\n" + s5.ToString());
                srOutput.Close();
                fsOutput.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "writeSchema");
            }
            finally
            { }
        }
        // Delimiter character selection
        private void Format()
        {
            try
            {

                if (rdbImportSemicolon.Checked)
                {
                    strFormat = "Delimited(;)";
                }
                else if (rdbImportTab.Checked)
                {
                    strFormat = "TabDelimited";
                }
                else if (rdbImportOther.Checked)
                {
                    strFormat = "Delimited(" + txtSeparatorOtherChar.Text.Trim() + ")";
                }
                else
                {
                    strFormat = "Delimited(;)";
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Format");
            }
            finally
            {
            }
        }

        public DataSet LoadCSV(int numberOfRows)
        {
            DataSet ds = new DataSet();
            try
            {
                // Creates and opens an ODBC connection
                string strConnString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + this.dirCSV.Trim() + ";Extensions=asc,csv,tab,txt;Persist Security Info=False";
                string sql_select;
                OdbcConnection conn;
                conn = new OdbcConnection(strConnString.Trim());
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }

                //Creates the select command text
                if (numberOfRows == -1)
                {
                    sql_select = "select * from [" + this.FileNevCSV.Trim() + "]";
                }
                else
                {
                    sql_select = "select top " + numberOfRows + " * from [" + this.FileNevCSV.Trim() + "]";
                }

                //Creates the data adapter
                OdbcDataAdapter obj_oledb_da = new OdbcDataAdapter(sql_select, conn);

                //Fills dataset with the records from CSV file
                obj_oledb_da.Fill(ds, "csv");

                //closes the connection
                conn.Close();
            }
            catch (Exception e) //Error
            {
                MessageBox.Show(e.Message, "Error - LoadCSV", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return ds;
        }

        private void txtFileToImport_TextChanged(object sender, EventArgs e)
        {
            // full file name
            this.fileCSV = this.txtFileToImport.Text;

            // creates a System.IO.FileInfo object to retrive information from selected file.
            System.IO.FileInfo fi = new System.IO.FileInfo(this.fileCSV);
            // retrives directory
            this.dirCSV = fi.DirectoryName.ToString();
            // retrives file name with extension
            this.FileNevCSV = fi.Name.ToString();

            // database table name
            this.txtTableName.Text = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length).Replace(" ", "_");
        }

        private void btnSave_DataSet_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);
            SaveToDatabase_withDataSet(conn);
        }

        // Checks if a file was given.
        private bool fileCheck()
        {
            if ((fileCSV == "") || (fileCSV == null) || (dirCSV == "") || (dirCSV == null) || (FileNevCSV == "") || (FileNevCSV == null))
            {
                MessageBox.Show("Select a CSV file to load first!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            else
            {
                return true;
            }
        }
        private long rowCount = 0;	//row number of source file
        private void SaveToDatabase_withDataSet(SqlConnection conn)
        {
            try
            {
                if (fileCheck())
                {
                    // select format, encoding, and write the schema file
                    Format();
                    ImportEncoding();
                    writeSchema();
                    if (conn.State.ToString() == "Closed")
                    {
                        conn.Open();
                    }
                    // loads all rows from from csv file
                    DataSet ds = LoadCSV(-1);

                    // gets the number of rows
                    this.rowCount = ds.Tables[0].Rows.Count;

                    // Makes a DataTableReader, which reads data from data set.
                    // It is nececery for bulk copy operation.
                    DataTableReader dtr = ds.Tables[0].CreateDataReader();

                    // Creates schema table. It gives column names for create table command.
                    DataTable dt;
                    dt = dtr.GetSchemaTable();

                    // You can view that schema table if you want:
                    //this.dataGridView_preView.DataSource = dt;

                    // Creates a new and empty table in the sql database
                    CreateTableInDatabase(dt, this.txtOwner.Text, this.txtTableName.Text, conn);

                    // Copies all rows to the database from the dataset.
                    using (SqlBulkCopy bc = new SqlBulkCopy(conn))
                    {
                        // Destination table with owner - this example doesn't
                        // check the owner and table names!
                        bc.DestinationTableName = "[" + this.txtOwner.Text + "].[" + this.txtTableName.Text + "]";

                        // User notification with the SqlRowsCopied event
                        bc.NotifyAfter = 100;
                        bc.SqlRowsCopied += new SqlRowsCopiedEventHandler(OnSqlRowsCopied);

                        // Starts the bulk copy.
                        bc.WriteToServer(ds.Tables[0]);

                        // Closes the SqlBulkCopy instance
                        bc.Close();
                    }

                    // Writes the number of imported rows to the form
                    this.lblProgress.Text = "Imported: " + this.rowCount.ToString() + "/" + this.rowCount.ToString() + " row(s)";
                    this.lblProgress.Refresh();

                    // Notifies user
                    MessageBox.Show("ready");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error - SaveToDatabase_withDataSet", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /*
     *  shows the progress of import operation
     */

        private void OnSqlRowsCopied(object sender, SqlRowsCopiedEventArgs e)
        {
            this.lblProgress.Text = "Imported: " + e.RowsCopied.ToString() + "/" + this.rowCount.ToString() + " row(s)";
            this.lblProgress.Refresh();
        }


        private bool CreateTableInDatabase(DataTable dtSchemaTable, string tableOwner, string tableName, SqlConnection conn)
        {
            try
            {
                if (conn.State.ToString() == "Closed")
                {
                    conn.Open();
                }
                // Generates the create table command.
                // The first column of schema table contains the column names.
                // The data type is nvarcher(4000) in all columns.

                string ctStr = "CREATE TABLE [" + tableOwner + "].[" + tableName + "](\r\n";
                for (int i = 0; i < dtSchemaTable.Rows.Count; i++)
                {
                    ctStr += "  [" + dtSchemaTable.Rows[i][0].ToString() + "] [nvarchar](4000) NULL";
                    if (i < dtSchemaTable.Rows.Count)
                    {
                        ctStr += ",";
                    }
                    ctStr += "\r\n";
                }
                ctStr += ")";

                // You can check the sql statement if you want:
                //MessageBox.Show(ctStr);


                // Runs the sql command to make the destination table.

             
                SqlCommand command = conn.CreateCommand();
                command.CommandText = ctStr;
                conn.Open();
                command.ExecuteNonQuery();
                conn.Close();

                return true;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "CreateTableInDatabase", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void btnSave_Direct_Click(object sender, EventArgs e)
        {
            var strServer = this.tbxServerName.Text;
            var strLogin = this.tbxLogin.Text;
            var strPassword = this.tbxPassword.Text;
            var strDbName = this.cbxDatabases.SelectedItem.ToString();

            var conn = Tools.GetConnectionString(strServer, strDbName, strLogin, strPassword);
            SaveToDatabaseDirectly(conn);
        }

        private void SaveToDatabaseDirectly(SqlConnection connsql)
        {
            try
            {
                if (fileCheck())
                {
                    // select format, encoding, and write the schema file
                    Format();
                    ImportEncoding();
                    writeSchema();
                    if (connsql.State.ToString() == "Closed")
                    {
                        connsql.Open();
                    }
                    // Creates and opens an ODBC connection
                    string strConnString = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" + this.dirCSV.Trim() + ";Extensions=asc,csv,tab,txt;Persist Security Info=False";
                    string sql_select;
                    OdbcConnection conn;
                    conn = new OdbcConnection(strConnString.Trim());
                    conn.Open();

                    //Counts the row number in csv file - with an sql query
                    OdbcCommand commandRowCount = new OdbcCommand("SELECT COUNT(*) FROM [" + this.FileNevCSV.Trim() + "]", conn);
                    this.rowCount = System.Convert.ToInt32(commandRowCount.ExecuteScalar());

                    // Creates the ODBC command
                    sql_select = "select * from [" + this.FileNevCSV.Trim() + "]";
                    OdbcCommand commandSourceData = new OdbcCommand(sql_select, conn);

                    // Makes on OdbcDataReader for reading data from CSV
                    OdbcDataReader dataReader = commandSourceData.ExecuteReader();

                    // Creates schema table. It gives column names for create table command.
                    DataTable dt;
                    dt = dataReader.GetSchemaTable();

                    // You can view that schema table if you want:
                    //this.dataGridView_preView.DataSource = dt;

                    // Creates a new and empty table in the sql database
                    CreateTableInDatabase(dt, this.txtOwner.Text, this.txtTableName.Text, connsql);

                    // Copies all rows to the database from the data reader.
                    using (SqlBulkCopy bc = new SqlBulkCopy("server=(local);database=Test_CSV_impex;Trusted_Connection=True"))
                    {
                        // Destination table with owner - this example doesn't
                        // check the owner and table names!
                        bc.DestinationTableName = "[" + this.txtOwner.Text + "].[" + this.txtTableName.Text + "]";

                        // User notification with the SqlRowsCopied event
                        bc.NotifyAfter = 100;
                        bc.SqlRowsCopied += new SqlRowsCopiedEventHandler(OnSqlRowsCopied);

                        // Starts the bulk copy.
                        bc.WriteToServer(dataReader);

                        // Closes the SqlBulkCopy instance
                        bc.Close();
                    }

                    // Writes the number of imported rows to the form
                    this.lblProgress.Text = "Imported: " + this.rowCount.ToString() + "/" + this.rowCount.ToString() + " row(s)";
                    this.lblProgress.Refresh();

                    // Notifies user
                    MessageBox.Show("ready");
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error - SaveToDatabaseDirectly", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
