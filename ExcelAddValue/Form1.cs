using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ExcelAddValue
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\alkan\Desktop\Movie_List.xls;Extended Properties='Excel 8.0; HDR=YES'");

        private void button1_Click(object sender, EventArgs e)
        {
            connection.Open();

            OleDbCommand cmd = new OleDbCommand("insert into [Sayfa1$] values(@p1,@p2,@p3,@p4,@p5)",connection);

            cmd.Parameters.AddWithValue("@p1", textBox1.Text);
            cmd.Parameters.AddWithValue("@p2", textBox2.Text);
            cmd.Parameters.AddWithValue("@p3", textBox3.Text);
            cmd.Parameters.AddWithValue("@p4", textBox4.Text);
            cmd.Parameters.AddWithValue("@p5", textBox5.Text);

            cmd.ExecuteNonQuery();
            connection.Close();

            MessageBox.Show("Kaydedildi !");


        }
    }
}
