using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {
        public string[,] array_zo;
        public Form3()
        {
            InitializeComponent();
        }

        //public void send_array(Array[] arr, DataGridView datagr)
        //{
        //    arr = array_zo;
        //    datagr.Columns.Add("Станок включен", "Станок включен"); // 11
        //    datagr.Columns.Add("Станок выключен", "Станок выключен"); // 12
        //    for (int i = 0; i < arr.Length; i++)
        //    {
        //        for (int j = 0; j < datagr.Rows.Count; j++)
        //        {
        //            if ((datagr[0, j].Value == arr[0]) & (datagr[1, j].Value == arr[1]))
        //            {
        //                datagr[11, j].Value = arr[15];
        //                datagr[12, j].Value = arr[16];
        //            }
                   
        //        }

        //    }
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            openFD_VipOper.Filter = "Отчёт WINNUM (*.txt)|*.txt";
            if (openFD_VipOper.ShowDialog() == DialogResult.OK) 
            {
                label2.Text = openFD_VipOper.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFD_ZagrOper.Filter = "Отчёт WINNUM (*.txt)|*.txt";
            if (openFD_ZagrOper.ShowDialog() == DialogResult.OK)
            {
                label3.Text = openFD_ZagrOper.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string result_zo, result1_zo;
            string line;
            
            StreamReader sr_zo = new StreamReader(openFD_ZagrOper.FileName);
            string file_zo = File.ReadAllText(openFD_ZagrOper.FileName);
            result_zo = file_zo.Replace(", ч.", "");
            result_zo = result_zo.Replace("\"", "");
            
            result_zo = result_zo.Replace("Оператор,Оборудование,Загрузка оператора,Ремонт,Аварийная остановка,Станок под нагрузкой,Программа выполняется,Другое,Нет калибров,Нет режущего инструмента,Поломка станка,Нет заготовки,Нет УП,Переналадка,Плановое техническое обслуживание,Станок включен,Станок выключен", "");
            result_zo = result_zo.Replace("/r/n", "");
            //textBox1.Text = result_zo;
            //result1_zo = Regex.Replace(result_zo, "/r/n", "");
            File.WriteAllText(@"zagr_oper.txt", result_zo);
            StreamReader reader_zo = new StreamReader(@"zagr_oper.txt");
            int k = 0;
            while ((line = reader_zo.ReadLine()) != null)
            {
                k++;
                string[] items = line.Split(',');
                if (items[0] != "")
                {   
                    array_zo[k,0] = items[0];
                    array_zo[k,1] = items[1];
                    array_zo[k,2] = items[15];
                    array_zo[k,3] = items[16];
                    //MessageBox.Show(array_zo[0] + " | " + array_zo[1] + " | " + array_zo[15] + " | " + array_zo[16]);
                    //textBox1.Text = array_zo[2];
                    //MessageBox.Show(array_zo[2]);
                    //DataRow row = dt.NewRow();
                    //items[3] = Regex.Replace(items[3], "\"", "");
                    //row["Оператор"] = items[3].Replace("\"", "");

                    //for (int i = 0; i < 100; i++)
                    //{
                    //    progressBar1.Value = i;
                    //}
                }
            }
            //Form1 f = new Form1(array_zo);
            //f.Show();
        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        
    }
}
