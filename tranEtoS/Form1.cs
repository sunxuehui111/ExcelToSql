using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Data.SqlClient;
namespace tranEtoS
{
    public partial class Form1 : Form
    {
        //private BackgroundWorker worker = new BackgroundWorker();
        //private System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        public Form1()
        {
            InitializeComponent(); 
        }

        //   void timer_Tick(object sender, EventArgs e)
        //{
        //    if (this.pgbWrite.Value < this.pgbWrite.Maximum)
        //        {
        //            this.pgbWrite.PerformStep();
        //        }
        //   }

        // void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        //{
        //     timer.Stop();
        //     this.pgbWrite.Value = this.pgbWrite.Maximum;
        //     MessageBox.Show(string.Format("将文件导入到数据库 {0}成功！", cbDataName.SelectedItem.ToString()));
        //}

        // void worker_DoWork(object sender, DoWorkEventArgs e)
        // {
        //     int count = 100;
        //    for (int i = 0; i < count; i++)
        //    {
        //    }
        //}
        public static string connString;
        public static string sendStr;
        public static string updateData;
        /// <summary>
        /// 导入  单击导入  第二次刷新 进度条控制
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string startStr = string.Format("开始向: {0} 数据库导入", cbDataName.SelectedItem.ToString());
            WriteLog(startStr);
                int m = 0, j = 0,counter = clbExcelSheet.Items.Count;
                for (; m < clbExcelSheet.Items.Count; m++)
                {
                    if (!clbExcelSheet.GetItemChecked(m))
                    {
                        j++;
                        counter -= 1;
                    }
                }
                this.pgbWrite.Value = 0;
                this.pgbWrite.Maximum = counter;
                this.pgbWrite.Step = 1;
            if (File.Exists(textBox2.Text))
            {
                if (m == j)
                {
                    MessageBox.Show("请选择Excel表！");
                    return;
                }
                else
                {
                    for (int i = 0; i < clbExcelSheet.Items.Count; i++)
                    {
                        if (clbExcelSheet.GetItemChecked(i))
                        {
                            TransferData(textBox2.Text, clbExcelSheet.GetItemText(clbExcelSheet.Items[i]), connString);
                            pgbWrite.Value++;
                            //timer.Interval = 100;
                            //timer.Tick += new EventHandler(timer_Tick);
                            //worker.WorkerReportsProgress = true;
                            //worker.DoWork += new DoWorkEventHandler(worker_DoWork);
                            //worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
                            //worker.RunWorkerAsync();
                            //timer.Start();
                        }
                    }
                    MessageBox.Show(string.Format("将文件导入到数据库 {0}成功！", cbDataName.SelectedItem.ToString()));
                    string ending = string.Format("结束向:{0} 数据库导入数据", cbDataName.SelectedItem.ToString());
                    WriteLog(ending);
                    initControl();
                }
            }
            else
            {
                MessageBox.Show("文件路径有误!");
            }
        }

        /// <summary>
        /// 浏览   获取路径名 将Excel表名加入checklistbox中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            this.clbExcelSheet.Items.Clear();
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            if (fd.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = fd.FileName;
                string strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + fd.FileName + ";" + "Extended Properties = 'Excel 8.0;HDR=Yes;IMEX=2';";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                DataTable sheetNames = conn.GetOleDbSchemaTable
(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                conn.Close();
                foreach (DataRow dr in sheetNames.Rows)
                {
                    string strSheetTableName = dr[2].ToString();
                    if (strSheetTableName.Contains("$") && strSheetTableName.Replace("'", "").EndsWith("$"))
                    {
                        strSheetTableName = strSheetTableName.Substring(0, strSheetTableName.Length - 1);//提取有效的sheet值
                    }
                    clbExcelSheet.Items.Add(strSheetTableName);
                }
            }
            clbExcelSheet.SetItemChecked(0, true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 选择Windows验证 改变
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbWindows_CheckedChanged(object sender, EventArgs e)
        {
            if (rbWindows.Checked)
            {
                UID.ReadOnly = true;
                Pwd.ReadOnly = true;
                UID.Text = "";
                Pwd.Text = "";
                this.cbDataName.Items.Clear();
                cbDataName.Items.Add("master");
                cbDataName.SelectedIndex = 0;
                connString = "server= " + textServer.Text + ";database= " + cbDataName.SelectedItem.ToString() + ";Trusted_Connection=SSPI";
            }
        }

        /// <summary>
        /// 选择Sql Server身份验证
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rbSql_CheckedChanged(object sender, EventArgs e)
        {
            if (rbSql.Checked)
            {
                UID.ReadOnly = false;
                Pwd.ReadOnly = false;
                UID.Text = "sa";
                Pwd.Text = "123";
                this.cbDataName.Items.Clear();
                cbDataName.Items.Add("master");
                cbDataName.SelectedIndex = 0;
                connString = string.Format("server = {0}; uid = {1}; pwd = {2}; database = {3}", textServer.Text, UID.Text, Pwd.Text, cbDataName.SelectedItem.ToString());
            }
        }

        /// <summary>
        /// 连接按钮
        /// 将数据库名加入Combobox中 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnrefresh_Click(object sender, EventArgs e)
        {
            if (cbDataName.SelectedItem.ToString() == "master")
            {
                this.cbDataName.Items.Clear();
                SqlConnection Connection = new SqlConnection(
                connString);
                DataTable DBNameTable = new DataTable();
                SqlDataAdapter Adapter = new SqlDataAdapter("select name from master..sysdatabases", Connection);

                lock (Adapter)
                {
                    Adapter.Fill(DBNameTable);
                }

                foreach (DataRow row in DBNameTable.Rows)
                {
                    cbDataName.Items.Add(row["name"]);
                    cbDataName.SelectedIndex = 0;
                }
                Connection.Close();
            }
        }

        /// <summary>
        /// combobox 选择项
        /// 将数据库中的表遍历打印在listbox中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbDataName_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbDataName.SelectedItem.ToString() != "master")
            {
                this.lbDataName.Items.Clear();
                if (rbWindows.Checked)
                    connString = "server= " + textServer.Text + ";database= " + cbDataName.SelectedItem.ToString() + ";Trusted_Connection=SSPI";
                else
                    connString = string.Format("server = {0}; uid = {1}; pwd = {2}; database = {3}", textServer.Text, UID.Text, Pwd.Text, cbDataName.SelectedItem.ToString());
                SqlConnection Connection = new SqlConnection(
                   connString);
                Connection.Open();
                //IList<string> tableName = new List<string>();
                //tableName = list(Connection);
                DataTable dataTable = Connection.GetSchema("Tables");
                foreach (DataRow row in dataTable.Rows)
                {
                    string tableType = (string)row["TABLE_TYPE"];
                    if (tableType.Contains("TABLE"))
                    {
                        lbDataName.Items.Add(row["TABLE_NAME"].ToString());
                    }
                }
                //for (int i = 0; i < tableName.Count; i++)
                //{
                //    lbDataName.Items.Add(tableName[i]);
                //}
                Connection.Close();
            }
        }

        /// <summary>
        /// 双击进入数据库表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lbDataName_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = this.lbDataName.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                sendStr = "select * from " + lbDataName.SelectedItem.ToString();
                Form2 f2 = new Form2();
                f2.ShowDialog();
            }
            else
                lbDataName.SelectedIndex = -1;
        }

        /// <summary>
        /// 关联Enter键
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Pwd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)//如果输入的是回车键  
            {
                this.btnrefresh_Click(sender, e);//触发button事件  
            }
        }
        //public IList<string> list(SqlConnection sqlCon)
        //{
        //    IList<string> tableList = new List<string>();
        //    DataTable dataTable = sqlCon.GetSchema("Tables");
        //    foreach (DataRow row in dataTable.Rows)
        //    {
        //        string tableType = (string)row["TABLE_TYPE"];
        //        if (tableType.Contains("TABLE"))
        //        {
        //            tableList.Add(row["TABLE_NAME"].ToString());
        //        }
        //    }
        //    return tableList;
        //}

        private void rbSql_MouseClick(object sender, MouseEventArgs e)
        {
            if (Pwd.Text != null)
            {
                this.btnrefresh_Click(sender, e);//触发button事件  
            }
        }

        private void rbWindows_MouseClick(object sender, MouseEventArgs e)
        {
            this.btnrefresh_Click(sender, e);//触发button事件  
        }


        public bool Exist(string tableName, string connectionString)
        {
            bool bExist = false;
            SqlConnection _Connection = new SqlConnection(connectionString);
            try
            {
                _Connection.Open();
                using (DataTable dt = _Connection.GetSchema("Tables"))
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string str1 = dr["TABLE_NAME"].ToString();
                    if (string.Equals(tableName, str1))
                    {
                                bExist = true;
                                break;
                    }
                    
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                _Connection.Dispose();
            }

            return bExist;
        }
        /// <summary>
        /// Excel写入数据库 实现函数
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="sheetName"></param>
        /// <param name="connectionString"></param>
        public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
                //获取全部数据     
                string strConn = "Provider = Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + excelFile + ";" + "Extended Properties = 'Excel 8.0;HDR=Yes;IMEX=2';";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}$]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(ds, sheetName);


                //如果目标表不存在则创建,excel文件的第一行为列标题,从第二行开始全部都是数据记录     
                string strSql = string.Format("if exists(select * from sysobjects where name = '{0}') truncate table {0}", sheetName);   //以sheetName为表名  
                
                if (!Exist(sheetName, connString))
                {
                    string ex = "数据库:" + sheetName + "表不存在";
                    WriteLog(ex);
                    return;
                }
                //foreach (System.Data.DataColumn c in ds.Tables[0].Columns)
                //{
                //    strSql += string.Format("[{0}] varchar(255),", c.ColumnName);
                //}
                //strSql = strSql.Trim(',') + ")";

                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
                //用bcp导入数据        
                //excel文件中列的顺序必须和数据表的列顺序一致，因为数据导入时，是从excel文件的第二行数据开始，不管数据表的结构是什么样的，反正就是第一列的数据会插入到数据表的第一列字段中，第二列的数据插入到数据表的第二列字段中，以此类推，它本身不会去判断要插入的数据是对应数据表中哪一个字段的     
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(connectionString))
                {
                    bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
                    bcp.BatchSize = 100;//每次传输的行数        
                    bcp.NotifyAfter = 100;//进度提示的行数        
                    bcp.DestinationTableName = sheetName;//目标表        
                    bcp.WriteToServer(ds.Tables[0]);
                }
                conn.Close();
                if (updateData == null)
                    updateData = "0";
                string success = string.Format("导入 {0} 表成功,更新 {1} 行", sheetName, updateData);
                WriteLog(success);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                WriteLog(ex.Message);
            }
        }
        //进度显示        
        void bcp_SqlRowsCopied(object sender, System.Data.SqlClient.SqlRowsCopiedEventArgs e)
        {
            this.Text = e.RowsCopied.ToString();
            updateData = e.RowsCopied.ToString();
            this.Update();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            connString = "";
            UID.Text = "";
            UID.ReadOnly = true;
            Pwd.Text = "";
            Pwd.ReadOnly = true;
            this.cbDataName.Items.Clear();
            rbSql.Checked = false;
            rbWindows.Checked = false;
            MessageBox.Show("连接已断开！");
        }
        public void initControl()
        {
            textBox2.Text = "";
            connString = "";
            UID.Text = "";
            UID.ReadOnly = true;
            Pwd.Text = "";
            Pwd.ReadOnly = true;
            this.cbDataName.Items.Clear();
            rbSql.Checked = false;
            rbWindows.Checked = false;
            clbExcelSheet.Items.Clear();
            this.pgbWrite.Value = 0;
            lbDataName.Items.Clear();
            this.Update();
        }

        private void cball_CheckedChanged(object sender, EventArgs e)
        {
            if (cball.Checked)
            {
                for (int j = 0; j < clbExcelSheet.Items.Count; j++)
                    clbExcelSheet.SetItemChecked(j, true);
            }
            else
            {
                for (int j = 0; j < clbExcelSheet.Items.Count; j++)
                    clbExcelSheet.SetItemChecked(j, false);
            }
        }


        public static void WriteLog(string strLog)
        {
            string sFilePath = "d:\\" + DateTime.Now.ToString("yyyyMM");
            string sFileName = "rizhi" + DateTime.Now.ToString("dd") + ".log";
            sFileName = sFilePath + "\\" + sFileName; //文件的绝对路径
            if (!Directory.Exists(sFilePath))//验证路径是否存在
            {
                Directory.CreateDirectory(sFilePath);
                //不存在则创建
            }
            FileStream fs;
            StreamWriter sw;
            if (File.Exists(sFileName))
            //验证文件是否存在，有则追加，无则创建
            {
                fs = new FileStream(sFileName, FileMode.Append, FileAccess.Write);
            }
            else
            {
                fs = new FileStream(sFileName, FileMode.Create, FileAccess.Write);
            }
            sw = new StreamWriter(fs);
            sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "   ---   " + strLog);
            sw.Close();
            fs.Close();
        }
    }
}