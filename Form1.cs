using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public string[,] array_zo;
        public string[,] array_ts;
        public string[,] dataar;
        string filename;
        private int ERR_WIN = 0;

        
        DataTable dt = new DataTable("Операторы");
        DataTable TS = new DataTable();
        DataTable ZO = new DataTable();
        public Form1()
        {
            InitializeComponent();

        }

        public int TotalLines(string filePath)
        {
            using (StreamReader r = new StreamReader(filePath))
            {
                int i = 0;
                while (r.ReadLine() != null) { i++; }
                return i;
            }
        }

        private void времяВСостоянииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (времяВСостоянииToolStripMenuItem.Checked == true)
            {
                времяВСостоянииToolStripMenuItem.Checked = false;
            }
            else
            {
                if (openFD_TimeS.ShowDialog() == DialogResult.OK)
                {
                    времяВСостоянииToolStripMenuItem.Checked = true;

                try
                {
                    TS.Columns.Add("Станок");
                    TS.Columns.Add("Статус");
                    TS.Columns.Add("Программа");
                    TS.Columns.Add("Начало");
                    TS.Columns.Add("Конец");
                        TS.Columns.Add("Длительность, ч.");
                    
                }
                catch { }
                    string result_ts, result1_ts;
                    string line;
                //    if (radioButton3.Checked == true)
                //{
                //    TS.Clear();
                //    dataGridView5.DataSource = null;
                //    dataGridView5.Rows.Clear();
                //} else if (radioButton4.Checked == true)
                    //{
                        //MessageBox.Show("Я здесь!");
                        openFD_TimeS.Multiselect = true;
                    openFD_TimeS.Filter = "Отчёт 'Время в состоянии' WINNUM (*.txt)|*.txt";
                    foreach (string fileWin in openFD_TimeS.FileNames)
                        {
                            
                            StreamReader sr_zo1 = new StreamReader(fileWin);
                            string file_zo1 = File.ReadAllText(fileWin);
                            result_ts = file_zo1.Replace(", ч.", "");
                            result_ts = result_ts.Replace("\"", "");
                            result_ts = result_ts.Replace("Fanuc серий T 30i,31i,32i,0iD,0iF", "");
                            result_ts = result_ts.Replace("Тип данных,Серийный номер,Действия,Наименование,Модель,Статус,Программа,Начальное значение,Последнее значение,Длительность (ч.)", "");
                            result_ts = result_ts.Replace("/r/n", "");
                            File.WriteAllText(@"times.txt", result_ts);
                            StreamReader reader_zo1 = new StreamReader(@"times.txt");
                            array_ts = new string[TotalLines(@"times.txt"), 6];
                            for (int k = 0; k < TotalLines(@"times.txt"); k++)
                            {
                                line = reader_zo1.ReadLine();
                                string[] items = line.Split(',');
                                if (items[0] != "")
                                {
                                    try
                                    {
                                        array_ts[k, 0] = items[3];
                                        array_ts[k, 1] = items[5];
                                        array_ts[k, 2] = items[6];
                                        array_ts[k, 3] = items[7].Substring(0, 19);
                                        array_ts[k, 4] = items[8].Substring(0, 19);
                                        string dlit = items[9] + '.' + items[10];
                                        array_ts[k, 5] = dlit;
                                    }
                                    catch { }
                                }
                            }
                        //MessageBox.Show(array_ts.Length.ToString());
                            for (int j = 0; j < TotalLines(@"times.txt"); j++)
                            {
                                DataRow TS_row = TS.NewRow();
                                for (int i = 0; i < 6; i++)
                                {
                                    if (array_ts[j, i] != null)
                                    {
                                        TS_row["Станок"] = array_ts[j, 0];
                                        TS_row["Статус"] = array_ts[j, 1];
                                        TS_row["Программа"] = array_ts[j, 2];
                                        TS_row["Начало"] = Convert.ToDateTime(array_ts[j, 3]);
                                        TS_row["Конец"] = Convert.ToDateTime(array_ts[j, 4]);
                                        TS_row["Длительность, ч."] = array_ts[j, 5];
                                    }
                                }
                                if (TS_row.IsNull(1) != true)
                                {

                                try
                                {
                                    TS.Rows.Add(TS_row);
                                }
                                catch { }
                                    
                                }
                                dataGridView5.DataSource = TS;
                            }
                        //MessageBox.Show(dataGridView5.Rows.Count.ToString());
                            reader_zo1.Close();
                        }
                }
            }
                
        }

        private void загрузкаОператоровToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (загрузкаОператоровToolStripMenuItem.Checked == true)
            {
                загрузкаОператоровToolStripMenuItem.Checked = false;
            }
            else
            {
                if (openFD_ZagrOper.ShowDialog() == DialogResult.OK)
                {
                    загрузкаОператоровToolStripMenuItem.Checked = true;
                try
                {
                    ZO.Columns.Add("ФИО");
                    ZO.Columns.Add("Станок");
                    ZO.Columns.Add("Станок включен");
                    ZO.Columns.Add("Станок выключен");
                }
                catch { }
                if (radioButton5.Checked == true)
                {
                    ZO.Clear();
                }
                string result_zo, result1_zo;
                string line;
                //comboBox1.Items.Clear();
                //comboBox1.Text= string.Empty;
                //comboBox2.Items.Clear();
                //comboBox2.Text = string.Empty;
                //dataGridView5.Rows.Clear();
                openFD_ZagrOper.Filter = "Отчёт WINNUM (*.txt)|*.txt";
                
                    StreamReader sr_zo = new StreamReader(openFD_ZagrOper.FileName);
                    string file_zo = File.ReadAllText(openFD_ZagrOper.FileName);
                    result_zo = file_zo.Replace(", ч.", "");
                    result_zo = result_zo.Replace("\"", "");

                    result_zo = result_zo.Replace("Оператор,Оборудование,Загрузка оператора,Аварийная остановка,Коррекция ускоренного хода 50%,Коррекция ускоренного хода 25%,Коррекция ускоренного хода 0%,Коррекция подачи выше 100%,Коррекция подачи ниже 100%,Коррекция скорости выше 100%,Коррекция скорости ниже 100%,Программа выполняется,Другое,Нет калибров,Нет режущего инструмента,Ремонт,Поломка станка,Нет заготовки,Нет УП,Переналадка,Плановое техническое обслуживание,Станок включен,Станок выключен", "");
                    result_zo = result_zo.Replace("/r/n", "");
                    //textBox1.Text = result_zo;
                    //result1_zo = Regex.Replace(result_zo, "/r/n", "");
                    File.WriteAllText(@"zagr_oper.txt", result_zo);
                    StreamReader reader_zo = new StreamReader(@"zagr_oper.txt");
                    ///array_zo = [reader_zo.ReadToEnd(), 4];  
                    int TotalLines(string filePath)
                    {
                        using (StreamReader r = new StreamReader(filePath))
                        {
                            int i = 0;
                            while (r.ReadLine() != null) { i++; }
                            return i;
                        }
                    }
                    MessageBox.Show(TotalLines(@"zagr_oper.txt").ToString());
                    array_zo = new string[TotalLines(@"zagr_oper.txt"), 4];
                    for (int k = 0; k < TotalLines(@"zagr_oper.txt"); k++)
                    {
                        line = reader_zo.ReadLine();
                        string[] items = line.Split(',');
                        if (items[0] != "")
                        {
                            try
                            {
                                array_zo[k, 0] = items[0];
                                //MessageBox.Show(items[1].ToString());
                                array_zo[k, 1] = items[1];
                                array_zo[k, 2] = items[21];
                                array_zo[k, 3] = items[22];
                            }
                            catch { }
                        }
                    }


                    //while ((line = reader_zo.ReadLine()) != null)
                    //{
                    //    k++;
                    //    string[] items = line.Split(',');
                    //    if (items[0] != "")
                    //    {
                    //        array_zo[k, 0] = items[0];
                    //        //MessageBox.Show(items[1].ToString());
                    //        array_zo[k, 1] = items[1];
                    //        array_zo[k, 2] = items[15];
                    //        array_zo[k, 3] = items[16];

                    //textBox1.Text = array_zo[2];
                    //MessageBox.Show(array_zo[2]);
                    //DataRow row = dt.NewRow();
                    //items[3] = Regex.Replace(items[3], "\"", "");
                    //row["Оператор"] = items[3].Replace("\"", "");

                    //for (int i = 0; i < 100; i++)
                    //{
                    //    progressBar1.Value = i;
                    //}
                    //    }

                    //}
                    //try {

                    for (int j = 0; j < TotalLines(@"zagr_oper.txt"); j++)
                    {
                        DataRow ZO_row = ZO.NewRow();
                        for (int i = 0; i < 4; i++)
                        {
                            //MessageBox.Show();
                            if (array_zo[j, i] != null)
                            {

                                ZO_row["ФИО"] = array_zo[j, 0];
                                ZO_row["Станок"] = array_zo[j, 1];
                                ZO_row["Станок включен"] = array_zo[j, 2];
                                ZO_row["Станок выключен"] = array_zo[j, 3];
                                //MessageBox.Show(array_zo[j, i].ToString());

                            }

                            //MessageBox.Show(array_zo[j,i] + " | " + array_zo[1] + " | " + array_zo[2] + " | " + array_zo[3]);
                        }
                        if (ZO_row.IsNull(1) != true)
                        {
                            ZO.Rows.Add(ZO_row);
                        }

                        dataGridView5.DataSource = ZO;
                    }

                    //} catch{}
                }
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string polez_b_h, polez_b_m, polez_b_s, ostanov_b_h, ostanov_b_m, ostanov_b_s, pause_b_h, pause_b_m, pause_b_s;
            int polez_s, ostanov_s, pause_s;

            openFD_VipOper.Filter = "Отчёт в TXT|*.txt";
            openFD_VipOper.DefaultExt = "*.txt";
            openFD_VipOper.FileName = "";

            if (выполненныеОперацииToolStripMenuItem.Checked == true)
            {
                выполненныеОперацииToolStripMenuItem.Checked = false;
            }
            else { 
            
                dt.Clear();
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                comboBox1.Text = string.Empty;
                comboBox2.Text = string.Empty;
                выполненныеОперацииToolStripMenuItem.Checked = true;
                if ((openFD_VipOper.ShowDialog() == DialogResult.OK))
                    //& (filename != openFD_VipOper.FileName))
                {
                    foreach (string fileWin in openFD_VipOper.FileNames)
                    {
                        выполненныеОперацииToolStripMenuItem.Checked = true;
                        
                        
                        dataGridView1.DataSource = null;
                        if (radioButton1.Checked == true)
                        {
                            dt.Clear();
                        }
                        dataGridView1.Rows.Clear();
                        filename = fileWin;
                        string result1;
                        string line;
                        string file = File.ReadAllText(fileWin);
                        string result = Regex.Replace(file, "\\[([^\\s]*)\\]", "");
                        result = result.Replace(",\"График выполнения программы [в новом окне]\"", "");
                        result = result.Replace("\"Имя программы\",Обозначение,Оборудование,Оператор,Начало,Завершение,Кадры,\"Завершенная операция\",\"Машинное время\",Остановы,\"Время изготовления\",Пауза", "");
                        result = result.Replace("\"", "");
                        result1 = Regex.Replace(result, "/r/n", "");

                        File.WriteAllText(@"result.txt", result1);

                        StreamReader reader1 = new StreamReader(@"result.txt");

                        while ((line = reader1.ReadLine()) != null)
                        {
                            string[] items = line.Split(',');
                            if (items[0] != "")
                            {
                                DataRow row = dt.NewRow();
                                items[3] = Regex.Replace(items[3], "\"", "");
                                row["Оператор"] = items[3].Replace("\"", "");
                                if (!comboBox1.Items.Contains(items[3].Replace("\"", "")))
                                {
                                    comboBox1.Items.Add(items[3].Replace("\"", ""));
                                    comboBox2.Items.Add(items[3].Replace("\"", ""));
                                }
                                row["Обозначение"] = items[1];
                                row["Программа"] = items[0];
                                row["Оборудование"] = items[2];
                                row["Начало"] = items[4].ToString().Replace("\"", "");
                                row["Завершение"] = items[5].ToString().Replace("\"", "");
                                row["Машинное время, ч."] = items[8].ToString().Replace("\"", "");
                                row["Остановы, ч."] = items[9].ToString().Replace("\"", "");
                                row["Пауза, ч."] = items[11].ToString().Replace("\"", "");
                                polez_b_h = items[8].ToString().Substring(0, 2);
                                polez_b_m = items[8].ToString().Substring(3, 2);
                                polez_b_s = items[8].ToString().Substring(6, 2);
                                polez_s = Convert.ToInt32(polez_b_s);
                                polez_s = polez_s + (Convert.ToInt32(polez_b_m) * 60);
                                polez_s = polez_s + (Convert.ToInt32(polez_b_h) * 3600);
                                row["Числитель, с."] = polez_s;
                                ostanov_b_h = items[9].ToString().Substring(0, 2);
                                ostanov_b_m = items[9].ToString().Substring(3, 2);
                                ostanov_b_s = items[9].ToString().Substring(6, 2);
                                ostanov_s = Convert.ToInt32(ostanov_b_s);
                                ostanov_s = ostanov_s + (Convert.ToInt32(ostanov_b_m) * 60);
                                ostanov_s = ostanov_s + (Convert.ToInt32(ostanov_b_h) * 3600);
                                if (items[11].Length > 0)
                                {
                                    pause_b_h = items[11].ToString().Substring(0, 2);
                                    pause_b_m = items[11].ToString().Substring(3, 2);
                                    pause_b_s = items[11].ToString().Substring(6, 2);
                                    try { pause_s = Convert.ToInt32(pause_b_s); }
                                    catch { MessageBox.Show("В выражении произошла ошибка на стороне WINNUM: '" + items[11] + "'"); reader1.Close(); ERR_WIN = 20; comboBox1.Items.Clear(); comboBox2.Items.Clear(); comboBox1.Text = ""; comboBox2.Text = ""; dataGridView2.DataSource = null; dataGridView2.Rows.Clear(); dataGridView3.DataSource = null; dataGridView3.Rows.Clear(); textBox1.Clear(); textBox2.Clear(); textBox3.Clear(); textBox10.Clear(); textBox11.Clear(); textBox12.Clear(); ; return; }
                                    pause_s = pause_s + (Convert.ToInt32(pause_b_m) * 60);
                                    pause_s = pause_s + (Convert.ToInt32(pause_b_h) * 3600);
                                    row["Знаменатель, с."] = pause_s + ostanov_s;
                                }
                                else
                                {
                                    pause_s = 0;
                                    row["Знаменатель, с."] = pause_s + ostanov_s;
                                }
                                dt.Rows.Add(row);
                            }
                        }
                        reader1.Close();

                        dataGridView1.DataSource = dt;
                        dataGridView1.Sort(dataGridView1.Columns[3], ListSortDirection.Ascending);
                    }
                    ERR_WIN = 0;
                }

            }
                
            
            //if (ERR_WIN == 20)
            //{
            //    reader1.Close();
            //    File.WriteAllText(@"result.txt", string.Empty);
            //    dataGridView1.DataSource = dt;
            //}
            //else
            //{

            //    dataGridView1.DataSource = dt;
            //    //reader.Close();
            //    //reader.Dispose();
            //    reader1.Close();
            //    ERR_WIN = 0;
            //}
            
        }

        private void repoday()
        {
            DataTable dt_repod = new DataTable();
            DataRow dr_repod = dt_repod.NewRow();
            dt_repod.Columns.Add("День");
            dt_repod.Columns.Add("Станок");
            dt_repod.Columns.Add("Машинное время");
            dt_repod.Columns.Add("Пауза");
            dt_repod.Columns.Add("Останов");
            dt_repod.Columns.Add("КПД, %");
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)
            //{
            //if (dataGridView4.Rows.Count > 0)
            //{
            //foreach (DataGridViewRow row in dataGridView1.Rows)
            //{
            //    string st = row.Cells["Начало"].ToString().Substring(0,10);
            //    if ()
            //    {

            //    }

            void finddg(string day)
            {
                for (int p = 0; p < 5; p++)
                {
                    try
                    {
                        if (!dataGridView4["День", p].Value.ToString().Contains(day))
                        {
                            MessageBox.Show(day);
                        }
                    }
                    catch { }
                }
            }
            //}

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                
                    string st = dataGridView1["Начало",i].ToString().Substring(0, 10);
                    //if (!dataGridView4.Rows[j].Cells["День"].Value.ToString().Contains(st))
                    //{
                    //    MessageBox.Show(st);
                    //} ;

                    finddg(st);
                
                
            }

                
                    //{ 
                    //    if (!row.Cells["День"].Value.ToString().Contains(dataGridView1["Начало", i].Value.ToString().Substring(0, 10)))
                    //    {
                    //        dr_repod["День"] = dataGridView1["Начало", i].Value.ToString().Substring(0, 10);
                            
                    //    }
                    //try
                    //{
                    //    dt_repod.Rows.Add(dr_repod);
                    //}
                    //catch { }

                    //}
                //}

               
                
            //}
            dataGridView4.DataSource = dt_repod;
        }

        private void comboras(string combo, DataGridView data, TextBox outp1, TextBox outp2, TextBox outp3, TextBox outp4, Label min, Label max, Label on, Label off)
        {
            int polez_s, all;
            double itog, pause_s, ostanov_s;
            string ostanov_b_h, ostanov_b_m, ostanov_b_s, pause_b_h, pause_b_m, pause_b_s;
            double[] pause_ar, ostanov_ar;
            string[] stanok;
            pause_s = 0;
            ostanov_s = 0;
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = dataGridView1.Columns[2].HeaderText.ToString() + " LIKE '%" + combo + "%'";
            polez_s = 0;
            all = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                polez_s += Convert.ToInt32(dataGridView1[9, i].Value);
                all += Convert.ToInt32(dataGridView1[10, i].Value);
            }
            itog = Math.Round(Convert.ToDouble(polez_s) / Convert.ToDouble((polez_s + all)) * 100, 2);
            outp4.Text = Convert.ToString(itog);
            ostanov_ar = new double[dataGridView1.RowCount];
            pause_ar = new double[dataGridView1.RowCount];
            outp1.Text = Convert.ToString(Math.Round(Convert.ToDouble(polez_s) / 60, 2));
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                ostanov_b_h = dataGridView1[7, j].Value.ToString().Substring(0, 2);
                ostanov_b_m = dataGridView1[7, j].Value.ToString().Substring(3, 2);
                ostanov_b_s = dataGridView1[7, j].Value.ToString().Substring(6, 2);
                ostanov_s = Convert.ToInt32(ostanov_b_s);
                ostanov_s = ostanov_s + (Convert.ToInt32(ostanov_b_m) * 60);
                ostanov_s = ostanov_s + (Convert.ToInt32(ostanov_b_h) * 3600);
                if (dataGridView1[8, j].Value.ToString().Length > 0)
                {
                    pause_b_h = dataGridView1[8, j].Value.ToString().Substring(0, 2);
                    pause_b_m = dataGridView1[8, j].Value.ToString().Substring(3, 2);
                    pause_b_s = dataGridView1[8, j].Value.ToString().Substring(6, 2);
                    pause_s = Convert.ToInt32(pause_b_s);
                    pause_s = pause_s + (Convert.ToInt32(pause_b_m) * 60);
                    pause_s = pause_s + (Convert.ToInt32(pause_b_h) * 3600);
                }
                ostanov_ar[j] = ostanov_s;
                pause_ar[j] = pause_s;
            }
            for (int y = 0; y < pause_ar.Count(); y++)
            {
                ostanov_s += Convert.ToInt32(ostanov_ar[y]);
                pause_s += Convert.ToInt32(pause_ar[y]);
            }
            outp2.Text = Convert.ToString(Math.Round(Convert.ToDouble(ostanov_s) / 60, 2));
            outp3.Text = Convert.ToString(Math.Round(Convert.ToDouble(pause_s) / 60, 2));
            stanok = new string[dataGridView1.Rows.Count];
            DataTable dt1 = new DataTable();
            DataRow dr1 = dt1.NewRow();
            data.Columns.Add("Станок", "Станок");
            if (загрузкаОператоровToolStripMenuItem.Checked == true)
            {
                //data.Columns.Add("Вкл.", "Вкл.");
                //data.Columns.Add("Выкл.", "Выкл.");
            }
            data.Columns.Add("Коэф.", "Коэф.");
            for (int y = 0; y < dataGridView1.Rows.Count; y++)
            {
                if (!listBox1.Items.Contains(dataGridView1["Оборудование", y].Value.ToString()))
                {
                    listBox1.Items.Add(dataGridView1["Оборудование", y].Value.ToString().Replace("\"", "").Trim());
                }
            }
            double ind = 0;
            for (int n = 0; n < listBox1.Items.Count; n++)
            {
                double all_e = 0;
                double a = 0;
                data.Rows.Add(listBox1.Items[n].ToString());
                for (int t = 0; t < dataGridView1.Rows.Count; t++)
                {
                    if (data[0, n].Value.ToString() == dataGridView1["Оборудование", t].Value.ToString())
                    {
                        all_e += (Convert.ToDouble(dataGridView1["Знаменатель, с.", t].Value) + Convert.ToDouble(dataGridView1["Числитель, с.", t].Value));
                        a += Convert.ToInt32(dataGridView1["Числитель, с.", t].Value);
                    }
                    ind = a / all_e * 100;
                    data["Коэф.", n].Value = Math.Round(ind, 2);
                }
            }
            var minValue_1 = Convert.ToDateTime("11.11.2011 00:00:00"); var maxValue_1 = Convert.ToDateTime("11.11.2011 00:00:00");
            try
            {
                var dateTimes = dataGridView1.Rows.Cast<DataGridViewRow>()
                .Select(x => Convert.ToDateTime(x.Cells["Начало"].Value));
                var minValue = dateTimes.Min();
                min.Text = minValue.ToString();
                var dateTimes1 = dataGridView1.Rows.Cast<DataGridViewRow>()
                .Select(y => Convert.ToDateTime(y.Cells["Завершение"].Value));
                var maxValue = dateTimes1.Max();
                max.Text = maxValue.ToString();
                minValue_1 = minValue; maxValue_1 = maxValue;
            }
            catch { MessageBox.Show("Обнаружены некорректные значения даты.");}
            double on1=0, off1 =0;
            
            //var maxValue = dateTimes.Max();
            //max.Text=maxValue.ToString();
            try { 
            for (int k = 0; k < dataGridView5.Rows.Count; k++)
            {
                for (int y = 0; y < data.Rows.Count; y++)
                {
                    if ((combo == dataGridView5[0, k].Value.ToString()) && (data[0,y].Value.ToString() == dataGridView5["Статус", k].Value.ToString()))
                    {
                        //on.Text = dataGridView5[2,k].Value.ToString();
                        data[1, y].Value = dataGridView5[2, k].Value.ToString();
                        //off.Text = dataGridView5[3, k].Value.ToString();
                        data[2, y].Value = dataGridView5[3, k].Value.ToString();
                        //MessageBox.Show(Convert.ToString(data[1, y].Value) + Convert.ToString(data[2, y].Value));
                        MessageBox.Show(data[1, y].Value.ToString());
                        on1 += Double.Parse(data[1, y].Value.ToString(), CultureInfo.InvariantCulture)*60;
                        off1 += Double.Parse(data[2, y].Value.ToString(), CultureInfo.InvariantCulture)*60;
                    }
                }
            }
            on.Text = on1.ToString();
            off.Text = off1.ToString();
                if (времяВСостоянииToolStripMenuItem.Checked == true)
                {
                    try
                    {
                        BindingSource bs_d5 = new BindingSource();
                        BindingSource bs_d5_1 = new BindingSource();
                        BindingSource bs_d5_2 = new BindingSource();
                        bs_d5.DataSource = dataGridView5.DataSource;
                        bs_d5_1.DataSource = dataGridView5.DataSource;
                        bs_d5_2.DataSource = dataGridView5.DataSource;
                        //try
                        //{
                        string f = "";
                        for (int h = 0; h < data.Rows.Count; h++)
                        {

                            f += dataGridView5.Columns["Станок"].HeaderText.ToString() + " LIKE '%" + data[0, h].Value.ToString() + "%' AND " + dataGridView5.Columns["Статус"].HeaderText.ToString() + " LIKE 'Станок выключен' AND Начало > '" + Convert.ToDateTime(minValue_1) + "' and Конец < '" + Convert.ToDateTime(maxValue_1) + "' OR ";
                            //bs_d5.Filter += 
                            //"AND " + 

                        }
                        int er = f.LastIndexOf("OR");
                        f = f.Substring(0, f.Length - 3);
                        f += "AND " + dataGridView5.Columns["Статус"].HeaderText.ToString() + " LIKE 'Станок выключен'";
                        //MessageBox.Show(f);
                        //f+= "AND " + dataGridView5.Columns[1].HeaderText.ToString() + " LIKE 'Станок выключен' AND Начало > '" + Convert.ToDateTime(minValue) + "' and Конец < '" + Convert.ToDateTime(maxValue) + "'";
                        //MessageBox.Show(f);
                        bs_d5.Filter = f.ToString();
                        //MessageBox.Show(bs_d5.Filter.ToString());
                        dataGridView5.DataSource = bs_d5;
                        dataGridView5.Sort(dataGridView1.Columns["Станок"], ListSortDirection.Ascending);
                        //}
                        //catch { }
                    }
                    catch { }
                }
            }
            catch { }
           

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            // string[] myArray;
            //MessageBox.Show(array_zo[1]);
        }
        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            //dataGridView5.DataSource = TS;
            comboras(comboBox1.Text, dataGridView2, textBox1, textBox2, textBox3, textBox6, label17, label18, label22, label23);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
            //dataGridView5.DataSource = TS;
            comboras(comboBox2.Text,dataGridView3,textBox12,textBox11,textBox10,textBox7, label20, label19,label25,label27);
            repoday();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter=null;
            dataGridView1.DataSource = bs;
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox6.Text = "";
            label17.Text = "";
            label18.Text = "";
            textBox6.Text = "";
            dataGridView2.Rows.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = dataGridView1.DataSource;
            bs.Filter = null;
            dataGridView1.DataSource = bs;
            textBox12.Text = "";
            textBox11.Text = "";
            textBox10.Text = "";
            textBox7.Text = "";
            label20.Text = "";
            label19.Text = "";
            dataGridView3.Rows.Clear();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dt.Columns.Add("Программа", typeof(String));
            dt.Columns.Add("Обозначение", typeof(String));
            dt.Columns.Add("Оператор", typeof(String));
            dt.Columns.Add("Оборудование", typeof(String));
            dt.Columns.Add("Начало", typeof(String));
            dt.Columns.Add("Завершение", typeof(String));
            dt.Columns.Add("Машинное время, ч.", typeof(String));
            dt.Columns.Add("Остановы, ч.", typeof(String));
            dt.Columns.Add("Пауза, ч.", typeof(String));
            dt.Columns.Add("Числитель, с.", typeof(int));
            dt.Columns.Add("Знаменатель, с.", typeof(int));
        }

        private void listBox1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
           
        }

        private void listBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            
        }

        private void listBox3_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            

        }
        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {

        }

      private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void label24_Click(object sender, EventArgs e)
        {

        }

        private void label23_Click(object sender, EventArgs e)
        {

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void отчётыToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
