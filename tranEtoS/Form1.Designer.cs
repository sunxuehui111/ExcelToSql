namespace tranEtoS
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.button1 = new System.Windows.Forms.Button();
            this.textServer = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cbDataName = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.UID = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.Pwd = new System.Windows.Forms.TextBox();
            this.btnrefresh = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lbDataName = new System.Windows.Forms.ListBox();
            this.rbSql = new System.Windows.Forms.RadioButton();
            this.rbWindows = new System.Windows.Forms.RadioButton();
            this.btnClose = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.pgbWrite = new System.Windows.Forms.ProgressBar();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.clbExcelSheet = new System.Windows.Forms.CheckedListBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(248, 473);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(135, 37);
            this.button1.TabIndex = 0;
            this.button1.Text = "导入";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textServer
            // 
            this.textServer.Location = new System.Drawing.Point(125, 29);
            this.textServer.Name = "textServer";
            this.textServer.Size = new System.Drawing.Size(201, 21);
            this.textServer.TabIndex = 1;
            this.textServer.Text = "(local)";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "数据库服务器名称：";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(316, 13);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(73, 24);
            this.button2.TabIndex = 0;
            this.button2.Text = "浏览";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(30, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(95, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "Excel文件地址：";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(131, 16);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(175, 21);
            this.textBox2.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 54);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 12);
            this.label3.TabIndex = 2;
            this.label3.Text = "需要操作Excel表：";
            // 
            // cbDataName
            // 
            this.cbDataName.FormattingEnabled = true;
            this.cbDataName.Location = new System.Drawing.Point(123, 173);
            this.cbDataName.Name = "cbDataName";
            this.cbDataName.Size = new System.Drawing.Size(203, 20);
            this.cbDataName.TabIndex = 3;
            this.cbDataName.SelectedIndexChanged += new System.EventHandler(this.cbDataName_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(84, 81);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 12);
            this.label4.TabIndex = 2;
            this.label4.Text = "ID：";
            // 
            // UID
            // 
            this.UID.Location = new System.Drawing.Point(122, 78);
            this.UID.Name = "UID";
            this.UID.ReadOnly = true;
            this.UID.Size = new System.Drawing.Size(204, 21);
            this.UID.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(72, 108);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 12);
            this.label5.TabIndex = 2;
            this.label5.Text = "密码：";
            // 
            // Pwd
            // 
            this.Pwd.Location = new System.Drawing.Point(122, 105);
            this.Pwd.Name = "Pwd";
            this.Pwd.PasswordChar = '*';
            this.Pwd.ReadOnly = true;
            this.Pwd.Size = new System.Drawing.Size(204, 21);
            this.Pwd.TabIndex = 5;
            this.Pwd.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Pwd_KeyDown);
            // 
            // btnrefresh
            // 
            this.btnrefresh.Location = new System.Drawing.Point(122, 132);
            this.btnrefresh.Name = "btnrefresh";
            this.btnrefresh.Size = new System.Drawing.Size(75, 24);
            this.btnrefresh.TabIndex = 0;
            this.btnrefresh.Text = "数据库连接";
            this.btnrefresh.UseVisualStyleBackColor = true;
            this.btnrefresh.Click += new System.EventHandler(this.btnrefresh_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lbDataName);
            this.groupBox1.Controls.Add(this.rbSql);
            this.groupBox1.Controls.Add(this.cbDataName);
            this.groupBox1.Controls.Add(this.rbWindows);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.Pwd);
            this.groupBox1.Controls.Add(this.textServer);
            this.groupBox1.Controls.Add(this.UID);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.btnClose);
            this.groupBox1.Controls.Add(this.btnrefresh);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Location = new System.Drawing.Point(12, 172);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(377, 295);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "服务器验证";
            // 
            // lbDataName
            // 
            this.lbDataName.FormattingEnabled = true;
            this.lbDataName.ItemHeight = 12;
            this.lbDataName.Location = new System.Drawing.Point(118, 213);
            this.lbDataName.Name = "lbDataName";
            this.lbDataName.Size = new System.Drawing.Size(208, 76);
            this.lbDataName.TabIndex = 7;
            this.lbDataName.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.lbDataName_MouseDoubleClick);
            // 
            // rbSql
            // 
            this.rbSql.AutoSize = true;
            this.rbSql.Location = new System.Drawing.Point(208, 56);
            this.rbSql.Name = "rbSql";
            this.rbSql.Size = new System.Drawing.Size(107, 16);
            this.rbSql.TabIndex = 6;
            this.rbSql.Text = "sql server验证";
            this.rbSql.UseVisualStyleBackColor = true;
            this.rbSql.CheckedChanged += new System.EventHandler(this.rbSql_CheckedChanged);
            this.rbSql.MouseClick += new System.Windows.Forms.MouseEventHandler(this.rbSql_MouseClick);
            // 
            // rbWindows
            // 
            this.rbWindows.AutoSize = true;
            this.rbWindows.Location = new System.Drawing.Point(59, 56);
            this.rbWindows.Name = "rbWindows";
            this.rbWindows.Size = new System.Drawing.Size(89, 16);
            this.rbWindows.TabIndex = 6;
            this.rbWindows.Text = "windows验证";
            this.rbWindows.UseVisualStyleBackColor = true;
            this.rbWindows.CheckedChanged += new System.EventHandler(this.rbWindows_CheckedChanged);
            this.rbWindows.MouseClick += new System.Windows.Forms.MouseEventHandler(this.rbWindows_MouseClick);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(236, 132);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 24);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "断开连接";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(23, 213);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(89, 12);
            this.label7.TabIndex = 2;
            this.label7.Text = "数据库表名称：";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(35, 173);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 12);
            this.label6.TabIndex = 2;
            this.label6.Text = "数据库名称：";
            // 
            // pgbWrite
            // 
            this.pgbWrite.Location = new System.Drawing.Point(12, 481);
            this.pgbWrite.Name = "pgbWrite";
            this.pgbWrite.Size = new System.Drawing.Size(230, 23);
            this.pgbWrite.TabIndex = 8;
            // 
            // clbExcelSheet
            // 
            this.clbExcelSheet.FormattingEnabled = true;
            this.clbExcelSheet.Location = new System.Drawing.Point(131, 54);
            this.clbExcelSheet.Name = "clbExcelSheet";
            this.clbExcelSheet.Size = new System.Drawing.Size(207, 116);
            this.clbExcelSheet.TabIndex = 9;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 522);
            this.Controls.Add(this.clbExcelSheet);
            this.Controls.Add(this.pgbWrite);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "excel导入工具";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textServer;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbDataName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox UID;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox Pwd;
        private System.Windows.Forms.Button btnrefresh;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rbSql;
        private System.Windows.Forms.RadioButton rbWindows;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListBox lbDataName;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ProgressBar pgbWrite;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.CheckedListBox clbExcelSheet;
    }
}

