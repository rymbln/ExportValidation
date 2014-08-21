﻿namespace ExportValidation
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tbxServerName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbxLogin = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tbxPassword = new System.Windows.Forms.TextBox();
            this.btnDatabases = new System.Windows.Forms.Button();
            this.cbxDatabases = new System.Windows.Forms.ComboBox();
            this.btnProcedures = new System.Windows.Forms.Button();
            this.cbxProcedures = new System.Windows.Forms.ComboBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.label4 = new System.Windows.Forms.Label();
            this.tbxOutputPath = new System.Windows.Forms.TextBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.tbxProjectName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button12 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.lbxProcedures = new System.Windows.Forms.ListBox();
            this.label6 = new System.Windows.Forms.Label();
            this.lbxViews = new System.Windows.Forms.ListBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.lbxTables = new System.Windows.Forms.ListBox();
            this.btnExportToCSV = new System.Windows.Forms.Button();
            this.lbl01 = new System.Windows.Forms.Label();
            this.gpbEncoding = new System.Windows.Forms.GroupBox();
            this.rdbUTF8 = new System.Windows.Forms.RadioButton();
            this.rdbUTF7 = new System.Windows.Forms.RadioButton();
            this.rdbASCII = new System.Windows.Forms.RadioButton();
            this.rdbUnicode = new System.Windows.Forms.RadioButton();
            this.gpbSeparator = new System.Windows.Forms.GroupBox();
            this.txtSeparatorOtherChar = new System.Windows.Forms.TextBox();
            this.rdbSeparatorOther = new System.Windows.Forms.RadioButton();
            this.rdbTab = new System.Windows.Forms.RadioButton();
            this.rdbSemicolon = new System.Windows.Forms.RadioButton();
            this.chkFirstRowColumnNames = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button14 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnSave_Direct = new System.Windows.Forms.Button();
            this.btnSave_DataSet = new System.Windows.Forms.Button();
            this.txtTableName = new System.Windows.Forms.TextBox();
            this.txtOwner = new System.Windows.Forms.TextBox();
            this.lblTableName = new System.Windows.Forms.Label();
            this.lblOwner = new System.Windows.Forms.Label();
            this.lblProgress = new System.Windows.Forms.Label();
            this.dataGridView_preView = new System.Windows.Forms.DataGridView();
            this.btnPreview = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.rdbImportOEM = new System.Windows.Forms.RadioButton();
            this.rdbImportUnicode = new System.Windows.Forms.RadioButton();
            this.rdbImportAnsi = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.rdbImportOther = new System.Windows.Forms.RadioButton();
            this.rdbImportTab = new System.Windows.Forms.RadioButton();
            this.rdbImportSemicolon = new System.Windows.Forms.RadioButton();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.txtFileToImport = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.button15 = new System.Windows.Forms.Button();
            this.maskedTextBox1 = new System.Windows.Forms.MaskedTextBox();
            this.maskedTextBox2 = new System.Windows.Forms.MaskedTextBox();
            this.button16 = new System.Windows.Forms.Button();
            this.rdbWin1251 = new System.Windows.Forms.RadioButton();
            this.statusStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.gpbEncoding.SuspendLayout();
            this.gpbSeparator.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_preView)).BeginInit();
            this.groupBox5.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbxServerName
            // 
            this.tbxServerName.Location = new System.Drawing.Point(12, 29);
            this.tbxServerName.Name = "tbxServerName";
            this.tbxServerName.Size = new System.Drawing.Size(144, 20);
            this.tbxServerName.TabIndex = 0;
            this.tbxServerName.Text = ".\\SQL2008";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Server";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(159, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Login";
            // 
            // tbxLogin
            // 
            this.tbxLogin.Location = new System.Drawing.Point(162, 29);
            this.tbxLogin.Name = "tbxLogin";
            this.tbxLogin.Size = new System.Drawing.Size(144, 20);
            this.tbxLogin.TabIndex = 3;
            this.tbxLogin.Text = "sa";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(309, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Password";
            // 
            // tbxPassword
            // 
            this.tbxPassword.Location = new System.Drawing.Point(312, 29);
            this.tbxPassword.Name = "tbxPassword";
            this.tbxPassword.Size = new System.Drawing.Size(144, 20);
            this.tbxPassword.TabIndex = 5;
            this.tbxPassword.Text = "asuseee";
            // 
            // btnDatabases
            // 
            this.btnDatabases.Location = new System.Drawing.Point(462, 27);
            this.btnDatabases.Name = "btnDatabases";
            this.btnDatabases.Size = new System.Drawing.Size(94, 23);
            this.btnDatabases.TabIndex = 6;
            this.btnDatabases.Text = "Get Databases";
            this.btnDatabases.UseVisualStyleBackColor = true;
            this.btnDatabases.Click += new System.EventHandler(this.button1_Click);
            // 
            // cbxDatabases
            // 
            this.cbxDatabases.FormattingEnabled = true;
            this.cbxDatabases.Location = new System.Drawing.Point(559, 29);
            this.cbxDatabases.Name = "cbxDatabases";
            this.cbxDatabases.Size = new System.Drawing.Size(166, 21);
            this.cbxDatabases.TabIndex = 7;
            this.cbxDatabases.SelectedIndexChanged += new System.EventHandler(this.cbxDatabases_SelectedIndexChanged);
            // 
            // btnProcedures
            // 
            this.btnProcedures.Location = new System.Drawing.Point(732, 29);
            this.btnProcedures.Name = "btnProcedures";
            this.btnProcedures.Size = new System.Drawing.Size(100, 23);
            this.btnProcedures.TabIndex = 8;
            this.btnProcedures.Text = "Get Procedures";
            this.btnProcedures.UseVisualStyleBackColor = true;
            this.btnProcedures.Click += new System.EventHandler(this.btnProcedures_Click);
            // 
            // cbxProcedures
            // 
            this.cbxProcedures.FormattingEnabled = true;
            this.cbxProcedures.Location = new System.Drawing.Point(838, 29);
            this.cbxProcedures.Name = "cbxProcedures";
            this.cbxProcedures.Size = new System.Drawing.Size(165, 21);
            this.cbxProcedures.TabIndex = 9;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 534);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1515, 22);
            this.statusStrip1.TabIndex = 10;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel1.Text = "toolStripStatusLabel1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(355, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Output Folder";
            // 
            // tbxOutputPath
            // 
            this.tbxOutputPath.Location = new System.Drawing.Point(358, 76);
            this.tbxOutputPath.Name = "tbxOutputPath";
            this.tbxOutputPath.Size = new System.Drawing.Size(340, 20);
            this.tbxOutputPath.TabIndex = 12;
            this.tbxOutputPath.Text = "C:\\Users\\rymbln\\Desktop\\Output\\PDF";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(7, 19);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(340, 23);
            this.btnGenerate.TabIndex = 14;
            this.btnGenerate.Text = "Generate PDF";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(7, 48);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(340, 23);
            this.button1.TabIndex = 15;
            this.button1.Text = "Generate Excel Validation";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(7, 162);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(340, 23);
            this.button2.TabIndex = 16;
            this.button2.Text = "Generate Word";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tbxProjectName
            // 
            this.tbxProjectName.Location = new System.Drawing.Point(12, 76);
            this.tbxProjectName.Name = "tbxProjectName";
            this.tbxProjectName.Size = new System.Drawing.Size(340, 20);
            this.tbxProjectName.TabIndex = 17;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 61);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "Project Name";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(7, 189);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(340, 23);
            this.button3.TabIndex = 19;
            this.button3.Text = "Generate Word Album";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(7, 75);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(340, 23);
            this.button4.TabIndex = 20;
            this.button4.Text = "Generate Excel Export";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(7, 133);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(340, 23);
            this.button6.TabIndex = 22;
            this.button6.Text = "Generate Queries In Format";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(7, 104);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(340, 23);
            this.button7.TabIndex = 23;
            this.button7.Text = "Generate CSV Export";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click_1);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(7, 218);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(340, 23);
            this.button8.TabIndex = 24;
            this.button8.Text = "Generate SAS Export";
            this.button8.UseVisualStyleBackColor = true;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(704, 76);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(75, 23);
            this.button5.TabIndex = 25;
            this.button5.Text = "Browse";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click_1);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button12);
            this.groupBox1.Controls.Add(this.button11);
            this.groupBox1.Controls.Add(this.button10);
            this.groupBox1.Controls.Add(this.button9);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.lbxProcedures);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.lbxViews);
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.lbxTables);
            this.groupBox1.Controls.Add(this.btnExportToCSV);
            this.groupBox1.Controls.Add(this.lbl01);
            this.groupBox1.Controls.Add(this.gpbEncoding);
            this.groupBox1.Controls.Add(this.gpbSeparator);
            this.groupBox1.Controls.Add(this.chkFirstRowColumnNames);
            this.groupBox1.Location = new System.Drawing.Point(382, 113);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(621, 343);
            this.groupBox1.TabIndex = 26;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Export To CSV";
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(180, 227);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(163, 26);
            this.button12.TabIndex = 42;
            this.button12.Text = "Select All";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(349, 227);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(163, 26);
            this.button11.TabIndex = 41;
            this.button11.Text = "Select All";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(11, 227);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(163, 26);
            this.button10.TabIndex = 40;
            this.button10.Text = "Select All";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(180, 256);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(163, 26);
            this.button9.TabIndex = 39;
            this.button9.Text = "Clear Selection";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(346, 19);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(85, 13);
            this.label7.TabIndex = 38;
            this.label7.Text = "PROCEDURES:";
            // 
            // lbxProcedures
            // 
            this.lbxProcedures.FormattingEnabled = true;
            this.lbxProcedures.HorizontalScrollbar = true;
            this.lbxProcedures.Location = new System.Drawing.Point(349, 38);
            this.lbxProcedures.Name = "lbxProcedures";
            this.lbxProcedures.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lbxProcedures.Size = new System.Drawing.Size(163, 186);
            this.lbxProcedures.TabIndex = 37;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(184, 19);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(45, 13);
            this.label6.TabIndex = 36;
            this.label6.Text = "VIEWS:";
            // 
            // lbxViews
            // 
            this.lbxViews.FormattingEnabled = true;
            this.lbxViews.HorizontalScrollbar = true;
            this.lbxViews.Location = new System.Drawing.Point(180, 38);
            this.lbxViews.Name = "lbxViews";
            this.lbxViews.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lbxViews.Size = new System.Drawing.Size(163, 186);
            this.lbxViews.TabIndex = 35;
            // 
            // btnRefresh
            // 
            this.btnRefresh.Location = new System.Drawing.Point(11, 256);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(163, 26);
            this.btnRefresh.TabIndex = 34;
            this.btnRefresh.Text = "Refresh list";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // lbxTables
            // 
            this.lbxTables.FormattingEnabled = true;
            this.lbxTables.HorizontalScrollbar = true;
            this.lbxTables.Location = new System.Drawing.Point(11, 38);
            this.lbxTables.Name = "lbxTables";
            this.lbxTables.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.lbxTables.Size = new System.Drawing.Size(163, 186);
            this.lbxTables.TabIndex = 29;
            // 
            // btnExportToCSV
            // 
            this.btnExportToCSV.Location = new System.Drawing.Point(11, 288);
            this.btnExportToCSV.Name = "btnExportToCSV";
            this.btnExportToCSV.Size = new System.Drawing.Size(597, 49);
            this.btnExportToCSV.TabIndex = 33;
            this.btnExportToCSV.Text = "Export to CSV";
            this.btnExportToCSV.UseVisualStyleBackColor = true;
            this.btnExportToCSV.Click += new System.EventHandler(this.btnExportToCSV_Click);
            // 
            // lbl01
            // 
            this.lbl01.AutoSize = true;
            this.lbl01.Location = new System.Drawing.Point(8, 22);
            this.lbl01.Name = "lbl01";
            this.lbl01.Size = new System.Drawing.Size(51, 13);
            this.lbl01.TabIndex = 28;
            this.lbl01.Text = "TABLES:";
            this.lbl01.Click += new System.EventHandler(this.lbl01_Click);
            // 
            // gpbEncoding
            // 
            this.gpbEncoding.Controls.Add(this.rdbWin1251);
            this.gpbEncoding.Controls.Add(this.rdbUTF8);
            this.gpbEncoding.Controls.Add(this.rdbUTF7);
            this.gpbEncoding.Controls.Add(this.rdbASCII);
            this.gpbEncoding.Controls.Add(this.rdbUnicode);
            this.gpbEncoding.Location = new System.Drawing.Point(518, 139);
            this.gpbEncoding.Name = "gpbEncoding";
            this.gpbEncoding.Size = new System.Drawing.Size(90, 140);
            this.gpbEncoding.TabIndex = 32;
            this.gpbEncoding.TabStop = false;
            this.gpbEncoding.Text = "Encoding";
            // 
            // rdbUTF8
            // 
            this.rdbUTF8.AutoSize = true;
            this.rdbUTF8.Location = new System.Drawing.Point(6, 88);
            this.rdbUTF8.Name = "rdbUTF8";
            this.rdbUTF8.Size = new System.Drawing.Size(52, 17);
            this.rdbUTF8.TabIndex = 4;
            this.rdbUTF8.Text = "UTF8";
            this.rdbUTF8.UseVisualStyleBackColor = true;
            // 
            // rdbUTF7
            // 
            this.rdbUTF7.AutoSize = true;
            this.rdbUTF7.Location = new System.Drawing.Point(6, 65);
            this.rdbUTF7.Name = "rdbUTF7";
            this.rdbUTF7.Size = new System.Drawing.Size(52, 17);
            this.rdbUTF7.TabIndex = 3;
            this.rdbUTF7.Text = "UTF7";
            this.rdbUTF7.UseVisualStyleBackColor = true;
            // 
            // rdbASCII
            // 
            this.rdbASCII.AutoSize = true;
            this.rdbASCII.Location = new System.Drawing.Point(6, 42);
            this.rdbASCII.Name = "rdbASCII";
            this.rdbASCII.Size = new System.Drawing.Size(52, 17);
            this.rdbASCII.TabIndex = 2;
            this.rdbASCII.Text = "ASCII";
            this.rdbASCII.UseVisualStyleBackColor = true;
            // 
            // rdbUnicode
            // 
            this.rdbUnicode.AutoSize = true;
            this.rdbUnicode.Location = new System.Drawing.Point(6, 19);
            this.rdbUnicode.Name = "rdbUnicode";
            this.rdbUnicode.Size = new System.Drawing.Size(65, 17);
            this.rdbUnicode.TabIndex = 1;
            this.rdbUnicode.Text = "Unicode";
            this.rdbUnicode.UseVisualStyleBackColor = true;
            // 
            // gpbSeparator
            // 
            this.gpbSeparator.Controls.Add(this.txtSeparatorOtherChar);
            this.gpbSeparator.Controls.Add(this.rdbSeparatorOther);
            this.gpbSeparator.Controls.Add(this.rdbTab);
            this.gpbSeparator.Controls.Add(this.rdbSemicolon);
            this.gpbSeparator.Location = new System.Drawing.Point(518, 32);
            this.gpbSeparator.Name = "gpbSeparator";
            this.gpbSeparator.Size = new System.Drawing.Size(90, 93);
            this.gpbSeparator.TabIndex = 30;
            this.gpbSeparator.TabStop = false;
            this.gpbSeparator.Text = "Separator";
            // 
            // txtSeparatorOtherChar
            // 
            this.txtSeparatorOtherChar.Location = new System.Drawing.Point(56, 65);
            this.txtSeparatorOtherChar.MaxLength = 1;
            this.txtSeparatorOtherChar.Name = "txtSeparatorOtherChar";
            this.txtSeparatorOtherChar.Size = new System.Drawing.Size(26, 20);
            this.txtSeparatorOtherChar.TabIndex = 3;
            // 
            // rdbSeparatorOther
            // 
            this.rdbSeparatorOther.AutoSize = true;
            this.rdbSeparatorOther.Location = new System.Drawing.Point(6, 65);
            this.rdbSeparatorOther.Name = "rdbSeparatorOther";
            this.rdbSeparatorOther.Size = new System.Drawing.Size(54, 17);
            this.rdbSeparatorOther.TabIndex = 2;
            this.rdbSeparatorOther.Text = "Other:";
            this.rdbSeparatorOther.UseVisualStyleBackColor = true;
            // 
            // rdbTab
            // 
            this.rdbTab.AutoSize = true;
            this.rdbTab.Location = new System.Drawing.Point(6, 42);
            this.rdbTab.Name = "rdbTab";
            this.rdbTab.Size = new System.Drawing.Size(46, 17);
            this.rdbTab.TabIndex = 1;
            this.rdbTab.Text = "TAB";
            this.rdbTab.UseVisualStyleBackColor = true;
            // 
            // rdbSemicolon
            // 
            this.rdbSemicolon.AutoSize = true;
            this.rdbSemicolon.Checked = true;
            this.rdbSemicolon.Location = new System.Drawing.Point(6, 19);
            this.rdbSemicolon.Name = "rdbSemicolon";
            this.rdbSemicolon.Size = new System.Drawing.Size(74, 17);
            this.rdbSemicolon.TabIndex = 0;
            this.rdbSemicolon.TabStop = true;
            this.rdbSemicolon.Text = "Semicolon";
            this.rdbSemicolon.UseVisualStyleBackColor = true;
            // 
            // chkFirstRowColumnNames
            // 
            this.chkFirstRowColumnNames.AutoSize = true;
            this.chkFirstRowColumnNames.Checked = true;
            this.chkFirstRowColumnNames.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkFirstRowColumnNames.Location = new System.Drawing.Point(350, 262);
            this.chkFirstRowColumnNames.Name = "chkFirstRowColumnNames";
            this.chkFirstRowColumnNames.Size = new System.Drawing.Size(156, 17);
            this.chkFirstRowColumnNames.TabIndex = 31;
            this.chkFirstRowColumnNames.Text = "First row has column names";
            this.chkFirstRowColumnNames.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button14);
            this.groupBox2.Controls.Add(this.button13);
            this.groupBox2.Controls.Add(this.btnGenerate);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.button8);
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.button7);
            this.groupBox2.Controls.Add(this.button4);
            this.groupBox2.Controls.Add(this.button6);
            this.groupBox2.Location = new System.Drawing.Point(15, 113);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(361, 308);
            this.groupBox2.TabIndex = 27;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Auto Export";
            // 
            // button14
            // 
            this.button14.Location = new System.Drawing.Point(7, 276);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(340, 23);
            this.button14.TabIndex = 26;
            this.button14.Text = "RUN SYNC";
            this.button14.UseVisualStyleBackColor = true;
            this.button14.Click += new System.EventHandler(this.button14_Click);
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(7, 248);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(340, 23);
            this.button13.TabIndex = 25;
            this.button13.Text = "RUN VALIDATION";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnSave_Direct);
            this.groupBox3.Controls.Add(this.btnSave_DataSet);
            this.groupBox3.Controls.Add(this.txtTableName);
            this.groupBox3.Controls.Add(this.txtOwner);
            this.groupBox3.Controls.Add(this.lblTableName);
            this.groupBox3.Controls.Add(this.lblOwner);
            this.groupBox3.Controls.Add(this.lblProgress);
            this.groupBox3.Controls.Add(this.dataGridView_preView);
            this.groupBox3.Controls.Add(this.btnPreview);
            this.groupBox3.Controls.Add(this.checkBox1);
            this.groupBox3.Controls.Add(this.groupBox5);
            this.groupBox3.Controls.Add(this.groupBox4);
            this.groupBox3.Controls.Add(this.btnBrowse);
            this.groupBox3.Controls.Add(this.txtFileToImport);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Location = new System.Drawing.Point(1020, 113);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(467, 343);
            this.groupBox3.TabIndex = 28;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Import From CSV";
            // 
            // btnSave_Direct
            // 
            this.btnSave_Direct.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSave_Direct.Location = new System.Drawing.Point(272, 309);
            this.btnSave_Direct.Name = "btnSave_Direct";
            this.btnSave_Direct.Size = new System.Drawing.Size(189, 28);
            this.btnSave_Direct.TabIndex = 20;
            this.btnSave_Direct.Text = "Save to database - directly";
            this.btnSave_Direct.UseVisualStyleBackColor = true;
            this.btnSave_Direct.Click += new System.EventHandler(this.btnSave_Direct_Click);
            // 
            // btnSave_DataSet
            // 
            this.btnSave_DataSet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSave_DataSet.Location = new System.Drawing.Point(272, 262);
            this.btnSave_DataSet.Name = "btnSave_DataSet";
            this.btnSave_DataSet.Size = new System.Drawing.Size(189, 28);
            this.btnSave_DataSet.TabIndex = 19;
            this.btnSave_DataSet.Text = "Save to database - with DataSet";
            this.btnSave_DataSet.UseVisualStyleBackColor = true;
            this.btnSave_DataSet.Click += new System.EventHandler(this.btnSave_DataSet_Click);
            // 
            // txtTableName
            // 
            this.txtTableName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtTableName.Location = new System.Drawing.Point(82, 317);
            this.txtTableName.Name = "txtTableName";
            this.txtTableName.Size = new System.Drawing.Size(164, 20);
            this.txtTableName.TabIndex = 18;
            // 
            // txtOwner
            // 
            this.txtOwner.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.txtOwner.Location = new System.Drawing.Point(82, 288);
            this.txtOwner.Name = "txtOwner";
            this.txtOwner.Size = new System.Drawing.Size(87, 20);
            this.txtOwner.TabIndex = 17;
            this.txtOwner.Text = "dbo";
            // 
            // lblTableName
            // 
            this.lblTableName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblTableName.AutoSize = true;
            this.lblTableName.Location = new System.Drawing.Point(6, 317);
            this.lblTableName.Name = "lblTableName";
            this.lblTableName.Size = new System.Drawing.Size(66, 13);
            this.lblTableName.TabIndex = 16;
            this.lblTableName.Text = "Table name:";
            // 
            // lblOwner
            // 
            this.lblOwner.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblOwner.AutoSize = true;
            this.lblOwner.Location = new System.Drawing.Point(6, 288);
            this.lblOwner.Name = "lblOwner";
            this.lblOwner.Size = new System.Drawing.Size(41, 13);
            this.lblOwner.TabIndex = 15;
            this.lblOwner.Text = "Owner:";
            // 
            // lblProgress
            // 
            this.lblProgress.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(6, 253);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(91, 13);
            this.lblProgress.TabIndex = 12;
            this.lblProgress.Text = "Imported: 0 row(s)";
            // 
            // dataGridView_preView
            // 
            this.dataGridView_preView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView_preView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_preView.Location = new System.Drawing.Point(9, 145);
            this.dataGridView_preView.Name = "dataGridView_preView";
            this.dataGridView_preView.Size = new System.Drawing.Size(452, 105);
            this.dataGridView_preView.TabIndex = 11;
            // 
            // btnPreview
            // 
            this.btnPreview.Location = new System.Drawing.Point(305, 114);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(156, 25);
            this.btnPreview.TabIndex = 10;
            this.btnPreview.Text = "Load preview (first 500 rows)";
            this.btnPreview.UseVisualStyleBackColor = true;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(305, 54);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(156, 17);
            this.checkBox1.TabIndex = 9;
            this.checkBox1.Text = "First row has column names";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.rdbImportOEM);
            this.groupBox5.Controls.Add(this.rdbImportUnicode);
            this.groupBox5.Controls.Add(this.rdbImportAnsi);
            this.groupBox5.Location = new System.Drawing.Point(144, 45);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(138, 94);
            this.groupBox5.TabIndex = 8;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Encoding";
            // 
            // rdbImportOEM
            // 
            this.rdbImportOEM.AutoSize = true;
            this.rdbImportOEM.Location = new System.Drawing.Point(6, 63);
            this.rdbImportOEM.Name = "rdbImportOEM";
            this.rdbImportOEM.Size = new System.Drawing.Size(49, 17);
            this.rdbImportOEM.TabIndex = 2;
            this.rdbImportOEM.Text = "OEM";
            this.rdbImportOEM.UseVisualStyleBackColor = true;
            // 
            // rdbImportUnicode
            // 
            this.rdbImportUnicode.AutoSize = true;
            this.rdbImportUnicode.Location = new System.Drawing.Point(6, 42);
            this.rdbImportUnicode.Name = "rdbImportUnicode";
            this.rdbImportUnicode.Size = new System.Drawing.Size(65, 17);
            this.rdbImportUnicode.TabIndex = 1;
            this.rdbImportUnicode.Text = "Unicode";
            this.rdbImportUnicode.UseVisualStyleBackColor = true;
            // 
            // rdbImportAnsi
            // 
            this.rdbImportAnsi.AutoSize = true;
            this.rdbImportAnsi.Checked = true;
            this.rdbImportAnsi.Location = new System.Drawing.Point(6, 19);
            this.rdbImportAnsi.Name = "rdbImportAnsi";
            this.rdbImportAnsi.Size = new System.Drawing.Size(50, 17);
            this.rdbImportAnsi.TabIndex = 0;
            this.rdbImportAnsi.TabStop = true;
            this.rdbImportAnsi.Text = "ANSI";
            this.rdbImportAnsi.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.textBox1);
            this.groupBox4.Controls.Add(this.rdbImportOther);
            this.groupBox4.Controls.Add(this.rdbImportTab);
            this.groupBox4.Controls.Add(this.rdbImportSemicolon);
            this.groupBox4.Location = new System.Drawing.Point(9, 45);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(129, 94);
            this.groupBox4.TabIndex = 6;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Separator";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(73, 66);
            this.textBox1.MaxLength = 1;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(24, 20);
            this.textBox1.TabIndex = 3;
            // 
            // rdbImportOther
            // 
            this.rdbImportOther.AutoSize = true;
            this.rdbImportOther.Location = new System.Drawing.Point(6, 65);
            this.rdbImportOther.Name = "rdbImportOther";
            this.rdbImportOther.Size = new System.Drawing.Size(54, 17);
            this.rdbImportOther.TabIndex = 2;
            this.rdbImportOther.Text = "Other:";
            this.rdbImportOther.UseVisualStyleBackColor = true;
            // 
            // rdbImportTab
            // 
            this.rdbImportTab.AutoSize = true;
            this.rdbImportTab.Location = new System.Drawing.Point(6, 42);
            this.rdbImportTab.Name = "rdbImportTab";
            this.rdbImportTab.Size = new System.Drawing.Size(46, 17);
            this.rdbImportTab.TabIndex = 1;
            this.rdbImportTab.Text = "TAB";
            this.rdbImportTab.UseVisualStyleBackColor = true;
            // 
            // rdbImportSemicolon
            // 
            this.rdbImportSemicolon.AutoSize = true;
            this.rdbImportSemicolon.Checked = true;
            this.rdbImportSemicolon.Location = new System.Drawing.Point(6, 19);
            this.rdbImportSemicolon.Name = "rdbImportSemicolon";
            this.rdbImportSemicolon.Size = new System.Drawing.Size(74, 17);
            this.rdbImportSemicolon.TabIndex = 0;
            this.rdbImportSemicolon.TabStop = true;
            this.rdbImportSemicolon.Text = "Semicolon";
            this.rdbImportSemicolon.UseVisualStyleBackColor = true;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(403, 17);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(58, 22);
            this.btnBrowse.TabIndex = 3;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // txtFileToImport
            // 
            this.txtFileToImport.Location = new System.Drawing.Point(94, 19);
            this.txtFileToImport.Name = "txtFileToImport";
            this.txtFileToImport.Size = new System.Drawing.Size(303, 20);
            this.txtFileToImport.TabIndex = 2;
            this.txtFileToImport.TextChanged += new System.EventHandler(this.txtFileToImport_TextChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 22);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(82, 13);
            this.label8.TabIndex = 1;
            this.label8.Text = "CSV file to load:";
            // 
            // button15
            // 
            this.button15.Location = new System.Drawing.Point(22, 425);
            this.button15.Name = "button15";
            this.button15.Size = new System.Drawing.Size(340, 23);
            this.button15.TabIndex = 27;
            this.button15.Text = "Activity Report";
            this.button15.UseVisualStyleBackColor = true;
            this.button15.Click += new System.EventHandler(this.button15_Click);
            // 
            // maskedTextBox1
            // 
            this.maskedTextBox1.Location = new System.Drawing.Point(22, 455);
            this.maskedTextBox1.Mask = "00/00/0000";
            this.maskedTextBox1.Name = "maskedTextBox1";
            this.maskedTextBox1.Size = new System.Drawing.Size(100, 20);
            this.maskedTextBox1.TabIndex = 29;
            // 
            // maskedTextBox2
            // 
            this.maskedTextBox2.Location = new System.Drawing.Point(262, 455);
            this.maskedTextBox2.Mask = "00/00/0000";
            this.maskedTextBox2.Name = "maskedTextBox2";
            this.maskedTextBox2.Size = new System.Drawing.Size(100, 20);
            this.maskedTextBox2.TabIndex = 30;
            this.maskedTextBox2.ValidatingType = typeof(System.DateTime);
            // 
            // button16
            // 
            this.button16.Location = new System.Drawing.Point(785, 75);
            this.button16.Name = "button16";
            this.button16.Size = new System.Drawing.Size(149, 23);
            this.button16.TabIndex = 31;
            this.button16.Text = "Open Folder";
            this.button16.UseVisualStyleBackColor = true;
            this.button16.Click += new System.EventHandler(this.button16_Click);
            // 
            // rdbWin1251
            // 
            this.rdbWin1251.AutoSize = true;
            this.rdbWin1251.Checked = true;
            this.rdbWin1251.Location = new System.Drawing.Point(6, 114);
            this.rdbWin1251.Name = "rdbWin1251";
            this.rdbWin1251.Size = new System.Drawing.Size(50, 17);
            this.rdbWin1251.TabIndex = 5;
            this.rdbWin1251.TabStop = true;
            this.rdbWin1251.Text = "ANSI";
            this.rdbWin1251.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1515, 556);
            this.Controls.Add(this.button16);
            this.Controls.Add(this.maskedTextBox2);
            this.Controls.Add(this.maskedTextBox1);
            this.Controls.Add(this.button15);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.tbxProjectName);
            this.Controls.Add(this.tbxOutputPath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.cbxProcedures);
            this.Controls.Add(this.btnProcedures);
            this.Controls.Add(this.cbxDatabases);
            this.Controls.Add(this.btnDatabases);
            this.Controls.Add(this.tbxPassword);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbxLogin);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbxServerName);
            this.Name = "Form1";
            this.Text = "Form1";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.gpbEncoding.ResumeLayout(false);
            this.gpbEncoding.PerformLayout();
            this.gpbSeparator.ResumeLayout(false);
            this.gpbSeparator.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_preView)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbxServerName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbxLogin;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbxPassword;
        private System.Windows.Forms.Button btnDatabases;
        private System.Windows.Forms.ComboBox cbxDatabases;
        private System.Windows.Forms.Button btnProcedures;
        private System.Windows.Forms.ComboBox cbxProcedures;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbxOutputPath;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox tbxProjectName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.ListBox lbxTables;
        private System.Windows.Forms.Button btnExportToCSV;
        private System.Windows.Forms.Label lbl01;
        private System.Windows.Forms.GroupBox gpbEncoding;
        private System.Windows.Forms.RadioButton rdbUTF8;
        private System.Windows.Forms.RadioButton rdbUTF7;
        private System.Windows.Forms.RadioButton rdbASCII;
        private System.Windows.Forms.RadioButton rdbUnicode;
        private System.Windows.Forms.GroupBox gpbSeparator;
        private System.Windows.Forms.TextBox txtSeparatorOtherChar;
        private System.Windows.Forms.RadioButton rdbSeparatorOther;
        private System.Windows.Forms.RadioButton rdbTab;
        private System.Windows.Forms.RadioButton rdbSemicolon;
        private System.Windows.Forms.CheckBox chkFirstRowColumnNames;
        private System.Windows.Forms.ListBox lbxViews;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ListBox lbxProcedures;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnSave_Direct;
        private System.Windows.Forms.Button btnSave_DataSet;
        private System.Windows.Forms.TextBox txtTableName;
        private System.Windows.Forms.TextBox txtOwner;
        private System.Windows.Forms.Label lblTableName;
        private System.Windows.Forms.Label lblOwner;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.DataGridView dataGridView_preView;
        private System.Windows.Forms.Button btnPreview;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.RadioButton rdbImportOEM;
        private System.Windows.Forms.RadioButton rdbImportUnicode;
        private System.Windows.Forms.RadioButton rdbImportAnsi;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.RadioButton rdbImportOther;
        private System.Windows.Forms.RadioButton rdbImportTab;
        private System.Windows.Forms.RadioButton rdbImportSemicolon;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.TextBox txtFileToImport;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button14;
        private System.Windows.Forms.Button button13;
        private System.Windows.Forms.Button button15;
        private System.Windows.Forms.MaskedTextBox maskedTextBox1;
        private System.Windows.Forms.MaskedTextBox maskedTextBox2;
        private System.Windows.Forms.Button button16;
        private System.Windows.Forms.RadioButton rdbWin1251;
    }
}

