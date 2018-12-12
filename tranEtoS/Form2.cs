using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace tranEtoS
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            str = Form1.sendStr;
            
        }
        private SqlConnection sqlCon;
        private string str;
        //public Form2(SqlConnection sqlCon1,string str1)
        //{
        //    InitializeComponent();
        //    sqlCon = sqlCon1;
        //    str = str1;
        //}

        private void Form2_Load(object sender, EventArgs e)
        {
            sqlCon = new SqlConnection(Form1.connString);
            sqlCon.Open();
            SqlCommand command = new SqlCommand();
            command.Connection = sqlCon;
            command.CommandText = str;
            command.ExecuteNonQuery();
            SqlDataAdapter sqldb = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            sqldb.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            sqlCon.Close();
        }
    }
}
