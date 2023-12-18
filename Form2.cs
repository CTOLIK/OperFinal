using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public String InboxData = String.Empty;
        public String s0,s1,s2,s3,s4,s5,s6,s7 = String.Empty;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            label1.Text = InboxData;
            label9.Text = s7;
            textBox1.Text = s0;
            textBox2.Text = s1;
            textBox3.Text = s2;
            textBox4.Text = s3;
            textBox5.Text = s4;
            textBox6.Text = s5;
            textBox7.Text = s6;
        }
    }
}
