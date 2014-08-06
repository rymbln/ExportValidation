using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
                var data = Tools.RunProcedure(conn, this.cbxProcedures.SelectedItem.ToString(), strProject);
                var index = Tools.GetIndex(conn, this.cbxProcedures.Text);
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
                var data = Tools.RunProcedure(conn, this.cbxProcedures.SelectedItem.ToString(), strProject);
                var index = Tools.GetIndex(conn, this.cbxProcedures.Text);
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
    }
}
