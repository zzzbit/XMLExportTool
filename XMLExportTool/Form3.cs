using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XMLExportTool
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form4.chargeName = textBox1.Text;
            Form4.chargeAddress = textBox3.Text;
            Form4.chargeCode = textBox4.Text;
            Form4.chargePhone = textBox5.Text;
            Form4.chargeEmail = textBox6.Text;
            Form4.visitLimit = textBox2.Text;
            Form4.classifyName = textBox7.Text;
            Form4.classifyVersion = textBox8.Text;
            this.Dispose();
        }
    }
}
