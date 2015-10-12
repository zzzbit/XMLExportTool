using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace XMLExportTool
{
    public partial class Form1 : Form
    {
        private DbOperation dboperation;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dboperation = new DbOperation();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool IsConnected = dboperation.Connect(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
            if (IsConnected)
            {
                Form4 from = new Form4(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text);
                from.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("登录失败");
            }
            dboperation.Close();
        }

    }
}
