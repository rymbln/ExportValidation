namespace ExportValidation
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
            this.button5 = new System.Windows.Forms.Button();
            this.statusStrip1.SuspendLayout();
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
            this.label2.Location = new System.Drawing.Point(9, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(33, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Login";
            // 
            // tbxLogin
            // 
            this.tbxLogin.Location = new System.Drawing.Point(12, 74);
            this.tbxLogin.Name = "tbxLogin";
            this.tbxLogin.Size = new System.Drawing.Size(144, 20);
            this.tbxLogin.TabIndex = 3;
            this.tbxLogin.Text = "sa";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 103);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Password";
            // 
            // tbxPassword
            // 
            this.tbxPassword.Location = new System.Drawing.Point(12, 119);
            this.tbxPassword.Name = "tbxPassword";
            this.tbxPassword.Size = new System.Drawing.Size(144, 20);
            this.tbxPassword.TabIndex = 5;
            this.tbxPassword.Text = "asuseee";
            // 
            // btnDatabases
            // 
            this.btnDatabases.Location = new System.Drawing.Point(12, 159);
            this.btnDatabases.Name = "btnDatabases";
            this.btnDatabases.Size = new System.Drawing.Size(144, 23);
            this.btnDatabases.TabIndex = 6;
            this.btnDatabases.Text = "Get Databases";
            this.btnDatabases.UseVisualStyleBackColor = true;
            this.btnDatabases.Click += new System.EventHandler(this.button1_Click);
            // 
            // cbxDatabases
            // 
            this.cbxDatabases.FormattingEnabled = true;
            this.cbxDatabases.Location = new System.Drawing.Point(12, 188);
            this.cbxDatabases.Name = "cbxDatabases";
            this.cbxDatabases.Size = new System.Drawing.Size(144, 21);
            this.cbxDatabases.TabIndex = 7;
            this.cbxDatabases.SelectedIndexChanged += new System.EventHandler(this.cbxDatabases_SelectedIndexChanged);
            // 
            // btnProcedures
            // 
            this.btnProcedures.Location = new System.Drawing.Point(12, 229);
            this.btnProcedures.Name = "btnProcedures";
            this.btnProcedures.Size = new System.Drawing.Size(144, 23);
            this.btnProcedures.TabIndex = 8;
            this.btnProcedures.Text = "Get Procedures";
            this.btnProcedures.UseVisualStyleBackColor = true;
            this.btnProcedures.Click += new System.EventHandler(this.btnProcedures_Click);
            // 
            // cbxProcedures
            // 
            this.cbxProcedures.FormattingEnabled = true;
            this.cbxProcedures.Location = new System.Drawing.Point(12, 259);
            this.cbxProcedures.Name = "cbxProcedures";
            this.cbxProcedures.Size = new System.Drawing.Size(144, 21);
            this.cbxProcedures.TabIndex = 9;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 348);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(535, 22);
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
            this.label4.Location = new System.Drawing.Point(180, 58);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Output Folder";
            // 
            // tbxOutputPath
            // 
            this.tbxOutputPath.Location = new System.Drawing.Point(183, 73);
            this.tbxOutputPath.Name = "tbxOutputPath";
            this.tbxOutputPath.Size = new System.Drawing.Size(340, 20);
            this.tbxOutputPath.TabIndex = 12;
            this.tbxOutputPath.Text = "C:\\Users\\rymbln\\Desktop\\Output\\PDF";
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(183, 119);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(340, 23);
            this.btnGenerate.TabIndex = 14;
            this.btnGenerate.Text = "Generate PDF";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(183, 159);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(340, 23);
            this.button1.TabIndex = 15;
            this.button1.Text = "Generate Excel Validation";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(183, 229);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(340, 23);
            this.button2.TabIndex = 16;
            this.button2.Text = "Generate Word";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // tbxProjectName
            // 
            this.tbxProjectName.Location = new System.Drawing.Point(183, 28);
            this.tbxProjectName.Name = "tbxProjectName";
            this.tbxProjectName.Size = new System.Drawing.Size(340, 20);
            this.tbxProjectName.TabIndex = 17;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(183, 13);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 13);
            this.label5.TabIndex = 18;
            this.label5.Text = "Project Name";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(183, 257);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(340, 23);
            this.button3.TabIndex = 19;
            this.button3.Text = "Generate Word Album";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(183, 186);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(340, 23);
            this.button4.TabIndex = 20;
            this.button4.Text = "Generate Excel Export";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(183, 298);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(340, 23);
            this.button5.TabIndex = 21;
            this.button5.Text = "Generate Queries";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(535, 370);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.tbxProjectName);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnGenerate);
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
        private System.Windows.Forms.Button button5;
    }
}

