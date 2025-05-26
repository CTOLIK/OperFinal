using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Management.Instrumentation;
using System.Linq.Expressions;
using Microsoft.VisualBasic;
using DocumentFormat.OpenXml.Drawing.Diagrams;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        //public static string connectString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\Files\обмен\Черненко А.А\База\ET\ET1.accdb";
        //private OleDbConnection myConnection=new OleDbConnection(connectString);

        public string[,] array_zo;
        public string[,] array_ts;
        //public DataTable dataar = new DataTable();
        //string filename;
        private int ERR_WIN = 0;
        //private DataGridView dataGridView1 = new DataGridView();
        private BindingSource bindingSource1 = new BindingSource();

        System.Data.DataTable dt = new System.Data.DataTable("Операторы");
        System.Data.DataTable TS = new System.Data.DataTable();
        System.Data.DataTable ZO = new System.Data.DataTable();
        List<string> legend = new List<string>();
        public Form1()
        {
            InitializeComponent();


        }

        private void id_rows_dv1(List<int> ird1)
        {
            
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
                                    array_ts[k, 4] = items[9].Substring(0, 19);
                                    string dlit = items[10] + '.' + items[11];
                                    array_ts[k, 5] = dlit;
                                }
                                catch { }
                            }
                        }
                    
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
                  
                    openFD_ZagrOper.Filter = "Отчёт WINNUM (*.txt)|*.txt";

                    StreamReader sr_zo = new StreamReader(openFD_ZagrOper.FileName);
                    string file_zo = File.ReadAllText(openFD_ZagrOper.FileName);
                    result_zo = file_zo.Replace(", ч.", "");
                    result_zo = result_zo.Replace("\"", "");

                    result_zo = result_zo.Replace("Оператор,Оборудование,Загрузка оператора,Аварийная Останов, ч.ка,Коррекция ускоренного хода 50%,Коррекция ускоренного хода 25%,Коррекция ускоренного хода 0%,Коррекция подачи выше 100%,Коррекция подачи ниже 100%,Коррекция скорости выше 100%,Коррекция скорости ниже 100%,Программа выполняется,Другое,Нет калибров,Нет режущего инструмента,Ремонт,Поломка станка,Нет заготовки,Нет УП,Переналадка,Плановое техническое обслуживание,Станок включен,Станок выключен", "");
                    result_zo = result_zo.Replace("/r/n", "");
                   
                    File.WriteAllText(@"zagr_oper.txt", result_zo);
                    StreamReader reader_zo = new StreamReader(@"zagr_oper.txt");
                   
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
                                array_zo[k, 1] = items[1];
                                array_zo[k, 2] = items[21];
                                array_zo[k, 3] = items[22];
                            }
                            catch { }
                        }
                    }


                    for (int j = 0; j < TotalLines(@"zagr_oper.txt"); j++)
                    {
                        DataRow ZO_row = ZO.NewRow();
                        for (int i = 0; i < 4; i++)
                        {
                          
                            if (array_zo[j, i] != null)
                            {

                                ZO_row["ФИО"] = array_zo[j, 0];
                                ZO_row["Станок"] = array_zo[j, 1];
                                ZO_row["Станок включен"] = array_zo[j, 2];
                                ZO_row["Станок выключен"] = array_zo[j, 3];
                             

                            }
                        }
                        if (ZO_row.IsNull(1) != true)
                        {
                            ZO.Rows.Add(ZO_row);
                        }

                        dataGridView5.DataSource = ZO;
                    }
                }
            }

        }

        public void button1_Click(object sender, EventArgs e)
        {

            string polez_b_h, polez_b_m, polez_b_s, ostanov_b_h, ostanov_b_m, ostanov_b_s, pause_b_h, pause_b_m, pause_b_s;
            int polez_s, ostanov_s, pause_s;
            progressBar1.Value = 0;
            openFD_VipOper.Filter = "Отчёт в TXT|*.txt";
            openFD_VipOper.DefaultExt = "*.txt";
            openFD_VipOper.FileName = "";
          

            if ((openFD_VipOper.ShowDialog() == DialogResult.OK))
        

            {
                dt.Clear();
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
              
                выполненныеОперацииToolStripMenuItem.Checked = true;
                foreach (string fileWin in openFD_VipOper.FileNames)
                {
                    выполненныеОперацииToolStripMenuItem.Checked = true;


                    dataGridView1.DataSource = null;
                    if (radioButton1.Checked == true)
                    {
                        dt.Clear();
                    }
                   
                    string result1;
                    string line;
                    string file = File.ReadAllText(fileWin);
                    string result = Regex.Replace(file, "\\[([^\\s]*)\\]", "");
                    result = result.Replace("\"Имя программы\",Обозначение,Оборудование,Оператор,Начало,Завершение,Кадры,Счетчик,\"Завершенная операция\",\"Машинное время\",Остановы,\"Время изготовления\",Пауза,\"График выполнения программы [в новом окне]\"", "");
                     
                    

                    result = result.Replace("\"", "");
                    result = result.Replace("KOL,CO", "");
                    result = result.Replace("KOL,CO", "");
                    result1 = Regex.Replace(result, "/r/n", "");

                    File.WriteAllText(@"result.txt", result1);

                    StreamReader reader1 = new StreamReader(@"result.txt");

                    int vs = 0;

                    while ((line = reader1.ReadLine()) != null)
                    {
                        vs++;
                        if (vs>1)
                        {
                            string[] items = line.Split(',');
                            if ((items[0] != "") && items.Count()>=14)
                            {
                                vs++;
                                DataRow row = dt.NewRow();
                                if (items[3] != "")
                                {
                                    items[3] = Regex.Replace(items[3], "\"", "");
                                    items[1] = Regex.Replace(items[1], "KOL,CO", "");
                                    items[1] = Regex.Replace(items[1], "KOL,CO", "");
                                    row["Оператор"] = items[3].Replace("\"", "");
                                    if (!comboBox1.Items.Contains(items[3].Replace("\"", "")))
                                    {
                                        comboBox1.Items.Add(items[3].Replace("\"", ""));
                                        comboBox2.Items.Add(items[3].Replace("\"", ""));
                                    }
                                }

                                row["Обозначение"] = items[1];
                                row["Программа"] = items[0];
                                row["Оборудование"] = items[2];
                                row["Начало"] = items[4].ToString().Replace("\"", "");
                                row["Завершение"] = items[5].ToString().Replace("\"", "");
                                row["Машинное время, ч."] = items[9].ToString().Replace("\"", "");
                                row["Норма времени"] = items[12].ToString().Replace("\"", "");
                                row["Остановы, ч."] = items[10].ToString().Replace("\"", "");
                                row["Счётчик"] = items[7].ToString().Replace("\"", "");
                                if (items[12].ToString().Length > 8)
                                {
                                    items[12] = items[12].ToString().Substring(1, 8);
                                }
                            
                                row["Пауза, ч."] = items[13].ToString().Replace("\"", "");
                                if (!items[9].ToString().Contains("-"))
                                {
                                    polez_b_h = items[9].ToString().Substring(0, 2);
                                    polez_b_m = items[9].ToString().Substring(3, 2);
                                    polez_b_s = items[9].ToString().Substring(6, 2);
                                    polez_s = Convert.ToInt32(polez_b_s);
                                    polez_s = polez_s + (Convert.ToInt32(polez_b_m) * 60);
                                    polez_s = polez_s + (Convert.ToInt32(polez_b_h) * 3600);
                                    try
                                    {
                                        row["Отклонение"] = TimeSpan.Parse(items[9].ToString().Replace("\"", "")) - TimeSpan.Parse(items[12].ToString().Replace("\"", ""));
                                    }
                                    catch { MessageBox.Show(items[9].ToString().Replace("\"", "") + " - " + items[12].ToString().Replace("\"", "")); };
                                   
                                    row["Числитель, с."] = polez_s;
                                   
                                    if (items[10].ToString() == string.Empty) { ostanov_s = 0; }
                                    else
                                    {
                                        ostanov_s = Convert.ToInt32(TimeSpan.Parse(items[10].ToString()).TotalSeconds);
                                    }

                                    if (items[13].Length > 0)
                                    {

                                        if (!items[13].ToString().Contains("-"))
                                        {
                                            pause_s = 0;
                                            if (items[13].ToString() == string.Empty) { pause_s = 0; } else
                                            {
                                                try { pause_s = Convert.ToInt32(TimeSpan.Parse(items[13].ToString()).TotalSeconds); }
                                                catch { pause_s = 0; MessageBox.Show("В некоторых данных Winnum ошибка вывода. Ошибочное значение в поле \"Пауза\": " + items[13].ToString()); }
                                                
                                            }
                                            row["Знаменатель, с."] = pause_s + ostanov_s;
                                        }
                                        else
                                        {
                                            pause_s = 0;
                                            pause_b_h = "0";
                                            pause_b_m = "0";
                                            pause_s = pause_s + (Convert.ToInt32(pause_b_m) * 60);
                                            pause_s = pause_s + (Convert.ToInt32(pause_b_h) * 3600);
                                            row["Знаменатель, с."] = pause_s + ostanov_s;
                                        }

                                    }
                                    else
                                    {
                                        pause_s = 0;
                                        row["Знаменатель, с."] = pause_s + ostanov_s;
                                    }
                                    dt.Rows.Add(row);
                                }
                            }
                        

                        }
                        comboBox1.Sorted = true;
                        comboBox2.Sorted = true;
                    }
                    reader1.Close();

                }
                ERR_WIN = 0;
            }
            DataView view = dt.DefaultView;
            view.Sort = "Оборудование ASC, Начало ASC";
            dataGridView1.DataSource = view;



            for (int u = 0; u < dataGridView1.Rows.Count; u++)
            {
                if (!legend.Contains(dataGridView1["Программа", u].Value.ToString() + " - " + dataGridView1["Обозначение", u].Value.ToString())) { legend.Add(dataGridView1["Программа", u].Value.ToString() + " - " + dataGridView1["Обозначение", u].Value.ToString()); }
            }

        }


        private void repoday(string operst)
        {
            listBox9.Items.Clear();
            System.Data.DataTable prostoi = new System.Data.DataTable();
            DataRow dr_prostoi = prostoi.NewRow();
            prostoi.Columns.Add("День");
            prostoi.Columns.Add("Станок");
            prostoi.Columns.Add("Программа");
            prostoi.Columns.Add("Простой");

            dataGridView3.Rows.Clear();
            dataGridView3.Columns.Clear();
           
            comboras(operst, dataGridView3, textBox12, textBox11, textBox10, textBox7, label20, label19, label25, label27);
           
            int count_smen = 0;
            List<string> id_start_days = new List<string>();
            List<string> id_end_days = new List<string>();
            List<string> stanok_list = new List<string>();

            id_start_days.Add(dataGridView1["Начало", 0].Value.ToString());
            stanok_list.Add(dataGridView1["Оборудование", 0].Value.ToString());
            TimeSpan stan = TimeSpan.Parse("-24:00:00");
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                DateTime day_end = Convert.ToDateTime(dataGridView1["Завершение", i].Value);
                TimeSpan period = Convert.ToDateTime(dataGridView1["Завершение", i + 1].Value) - day_end;
                string stanok = dataGridView1["Оборудование", i].Value.ToString();
                if ((period > Convert.ToDateTime("11.11.2011 04:00:00") - Convert.ToDateTime("11.11.2011 00:00:00")) || (stanok != dataGridView1["Оборудование", i + 1].Value.ToString()))
                {
                    // dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    count_smen++;
                }

            }
            
            int begin = 0;
            bool moreone = false;
            for (int i = begin; i < dataGridView1.Rows.Count - 1; i++)
            {
                string itog_period, start_day, end_day;

                DateTime day_end = Convert.ToDateTime(dataGridView1["Завершение", i].Value);
                DateTime smena_end;
                DateTime smena_start;
                TimeSpan period = Convert.ToDateTime(dataGridView1["Завершение", i + 1].Value) - day_end;
                string min_out = "";
                string stanok = dataGridView1["Оборудование", i].Value.ToString();

                if ((period > TimeSpan.Parse("04:00:00")) || (stanok != dataGridView1["Оборудование", i + 1].Value.ToString()))// || (i==0))
                {   
                    moreone=true;
                    //MessageBox.Show("1");
                    if (stanok != dataGridView1["Оборудование", i + 1].Value.ToString())
                    {
                        smena_end = day_end;
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Blue;
                        id_end_days.Add(dataGridView1["Завершение", i].Value.ToString());
                        id_start_days.Add(dataGridView1["Начало", i + 1].Value.ToString());
                        stanok_list.Add(dataGridView1["Оборудование", i + 1].Value.ToString());
                        //    }
                        //}
                        //MessageBox.Show("2");
                    }
                    else
                    {
                        //l1.Add(i);
                        smena_end = day_end;
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.Green;
                        id_end_days.Add(dataGridView1["Завершение", i].Value.ToString());
                        id_start_days.Add(dataGridView1["Начало", i + 1].Value.ToString());
                        stanok_list.Add(dataGridView1["Оборудование", i].Value.ToString());
                        //MessageBox.Show("3");
                    }
                }
               
            }

            TimeSpan ts_itog_n = TimeSpan.Zero;
            bool bfirst = false;
            List<string> Lnorm = new List<string>();
            List<string> Lmaxmv = new List<string>();
            TimeSpan maxmv = TimeSpan.Zero;

            for (int p = 0; p < dataGridView1.Rows.Count - 1; p++)
            {
                if ((p < dataGridView1.Rows.Count - 2) && (dataGridView1["Программа", p].Value.ToString() != dataGridView1["Программа", p + 1].Value.ToString()) && (dataGridView1["Программа", p + 1].Value.ToString() != dataGridView1["Программа", p + 2].Value.ToString()))
                {
                    TimeSpan ts_otkl_2 = TimeSpan.Parse(dataGridView1["Машинное время, ч.", p + 1].Value.ToString()) - TimeSpan.Parse(dataGridView1["Норма времени", p + 1].Value.ToString());
                    ts_itog_n += ts_otkl_2;
                    Lnorm.Add(dataGridView1["Оборудование", p + 1].Value.ToString() + ", " + dataGridView1["Программа", p + 1].Value.ToString() + ", " + dataGridView1["Начало", p + 1].Value.ToString() + ", " + dataGridView1["Завершение", p + 1].Value.ToString() + ", " + ts_otkl_2 + ", " + ts_otkl_2 + ", " + ts_otkl_2);
                }
                else
                if ((dataGridView1["Программа", p].Value.ToString() == dataGridView1["Программа", p + 1].Value.ToString()) && (dataGridView1["Программа", p].Value.ToString() != "O0"))
                {
                    if (dataGridView1["Оборудование", p].Value.ToString() == dataGridView1["Оборудование", p + 1].Value.ToString())
                    {
                        if (maxmv < TimeSpan.Parse(dataGridView1["Машинное время, ч.", p + 1].Value.ToString())) { maxmv = TimeSpan.Parse(dataGridView1["Машинное время, ч.", p + 1].Value.ToString()); }

                        bfirst = false;
                        TimeSpan ts_otkl = TimeSpan.Parse(dataGridView1["Машинное время, ч.", p + 1].Value.ToString()) - TimeSpan.Parse(dataGridView1["Норма времени", p + 1].Value.ToString());
                        ts_itog_n += ts_otkl;
                        Lnorm.Add(dataGridView1["Оборудование", p + 1].Value.ToString() + ", " + dataGridView1["Программа", p + 1].Value.ToString() + ", " + dataGridView1["Начало", p + 1].Value.ToString() + ", " + dataGridView1["Завершение", p + 1].Value.ToString() + ", " + ts_otkl + ", " + ts_itog_n + ", " + maxmv);

                    }



                }
                else {; ts_itog_n = TimeSpan.Zero; bfirst = true; maxmv = TimeSpan.Zero; Lmaxmv.Add(maxmv.ToString()); }
            }

            for (int l = 0; l < Lnorm.Count; l++)
            {
                listBox9.Items.Add(Lnorm[l].ToString());
            }

            id_end_days.Add(dataGridView1["Завершение", dataGridView1.Rows.Count - 1].Value.ToString());

            System.Data.DataTable dt_repod = new System.Data.DataTable();
            dt_repod.Columns.Add("Оператор");
            dt_repod.Columns.Add("Станок");
            dt_repod.Columns.Add("День");
            // dt_repod.Columns.Add("Смена");
            dt_repod.Columns.Add("Период", typeof(string));
            dt_repod.Columns.Add("Длительность смены, ч.", typeof(string));
            dt_repod.Columns.Add("Машинное время, ч.");
            dt_repod.Columns.Add("Отклонения", typeof(string));
            dt_repod.Columns.Add("Остановы, ч.");
            dt_repod.Columns.Add("Пауза, ч.");
            dt_repod.Columns.Add("КПД, %");
            dt_repod.Columns.Add("Номера программ");
           // dt_repod.Columns.Add("Инструмент");

            for (int i = 0; i < id_end_days.Count; i++)
            {
                if (Convert.ToDateTime(id_end_days[i]) > Convert.ToDateTime(id_start_days[i]))
                {
                    //MessageBox.Show(id_start_days[i].ToString() + " - " + id_end_days[i].ToString());
                    dataGridView4.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                    DataRow dr_repod = dt_repod.NewRow();
                    dr_repod["Оператор"] = operst;
                    dr_repod["Станок"] = stanok_list[i].ToString();
     //               day_list.Add(Convert.ToDateTime(id_start_days[i].ToString().Substring(0, 10)));
                    dr_repod["День"] = id_start_days[i].ToString().Substring(0, 10);
                    // dr_repod["Смена"] ="0";
                    dr_repod["Период"] = id_start_days[i].ToString() + " - " + id_end_days[i].ToString();
                    dr_repod["Длительность смены, ч."] = Convert.ToDateTime(id_end_days[i].ToString()) - Convert.ToDateTime(id_start_days[i].ToString());
                   // dr_repod["Номера программ"] = programs;
                    dt_repod.Rows.Add(dr_repod);
                }
            }

            dataGridView4.DataSource = dt_repod;

         //  myConnection.Open();
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                List<double> pause_ar = new List<double>();
                List<double> ostanov_ar = new List<double>();
                List<double> mash_ar = new List<double>();

                List<string> num_progs = new List<string>();
                List<string> num_progs2 = new List<string>();
                List<int> count = new List<int>();

                double itog, pause_s, ostanov_s, mash_time;
                string ostanov_b_h, ostanov_b_m, ostanov_b_s, pause_b_h, pause_b_m, pause_b_s, mash_s, mash_m, mash_h;
                pause_s = 0;
                ostanov_s = 0;
                mash_time = 0;

                string programs = string.Empty;

                List<int> prog = new List<int>();

                for (int t = 0; t < dataGridView1.Rows.Count; t++)
                {

                    if ((dataGridView4["Станок", i].Value.ToString() == dataGridView1["Оборудование", t].Value.ToString()) && (Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(0, 19)) <= Convert.ToDateTime(dataGridView1["Начало", t].Value.ToString())) && (Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(22, 19)) >= Convert.ToDateTime(dataGridView1["Завершение", t].Value.ToString())))
                    {
                        ostanov_s = Math.Round(Convert.ToDouble(TimeSpan.Parse(dataGridView1["Остановы, ч.", t].Value.ToString()).TotalSeconds) / 3600,2);
                   
                        if (!num_progs.Contains(dataGridView1["Программа", t].Value.ToString()))
                        {
                           // MessageBox.Show(dataGridView1["Программа", t].Value.ToString());
                            num_progs.Add(dataGridView1["Программа", t].Value.ToString());
                        }

                        if (ostanov_s >= 1)
                        {
                            dataGridView1["Остановы, ч.", t].Style.ForeColor = Color.Red;
                            dataGridView1["Остановы, ч.", t].Style.BackColor = Color.Black;
                        }
                        if (dataGridView1["Пауза, ч.", t].Value.ToString().Length > 0)
                        {
                                if (Convert.ToInt16(dataGridView1["Пауза, ч.", t].Value.ToString().Substring(0, 2)) <= 24)
                                {
                                    try
                                    {
                                        if (TimeSpan.Parse(dataGridView1["Пауза, ч.", t].Value.ToString()) < TimeSpan.Parse("5:00:00"))// && ()
                                        {
                                            pause_s = Math.Round(Convert.ToDouble(TimeSpan.Parse(dataGridView1["Пауза, ч.", t].Value.ToString()).TotalSeconds) / 3600, 2); 
                                        }
                                        if (Convert.ToDouble(pause_s) >= 1)
                                        {
                                            dataGridView1["Пауза, ч.", t].Style.ForeColor = Color.Red;
                                            dataGridView1["Пауза, ч.", t].Style.BackColor = Color.Black;
                                        }
                                    }
                                    catch { }
                                }
                        }
                        mash_h = dataGridView1["Машинное время, ч.", t].Value.ToString().Substring(0, 2);
                        mash_m = dataGridView1["Машинное время, ч.", t].Value.ToString().Substring(3, 2);
                        mash_s = dataGridView1["Машинное время, ч.", t].Value.ToString().Substring(6, 2);
                        mash_time = Convert.ToInt32(mash_s);
                        mash_time = mash_time + (Convert.ToInt32(mash_m) * 60);
                        mash_time = mash_time + (Convert.ToInt32(mash_h) * 3600);
                        ostanov_ar.Add(ostanov_s);
                        pause_ar.Add(pause_s);
                        mash_ar.Add(mash_time);

                    }

                }

                ostanov_s = 0;

                pause_s = 0;
                mash_time = 0;

                string day_prog = string.Empty;
                
                for (int y = 0; y < num_progs.Count; y++)
                {
                    string numofprog = num_progs[y].ToString();
                    string allpr = string.Empty;
                    TimeSpan tsallpr = TimeSpan.Zero;
                    
                    day_prog += "; "+numofprog + " - ";
                    int k = 0;
                    for (int o = 0; o < listBox9.Items.Count; o++)
                    {
                        if ((dataGridView4["Станок",i].Value.ToString() == listBox9.Items[o].ToString().Split(',')[0].Trim()) && (Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(0, 19)) <= Convert.ToDateTime(listBox9.Items[o].ToString().Split(',')[2].Trim())) && (Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(22, 19)) >= Convert.ToDateTime(listBox9.Items[o].ToString().Split(',')[3].Trim())))
                        {
                            if (num_progs[y].ToString() == listBox9.Items[o].ToString().Split(',')[1].Trim())
                            {
                                k++;

                                if (tsallpr < TimeSpan.Parse(listBox9.Items[o].ToString().Split(',')[4].Trim())) { tsallpr = TimeSpan.Parse(listBox9.Items[o].ToString().Split(',')[4].Trim()); allpr = tsallpr.ToString(); }
                               // allpr += listBox9.Items[o].ToString().Split(',')[4].Trim() + ", ";
                            }

                        // 

                        }
                    }

                   

                    TimeSpan sum_otkl = TimeSpan.Zero;

                    for (int o = 0; o < listBox9.Items.Count; o++)
                    {
                        if ((dataGridView4["Станок", i].Value.ToString() == listBox9.Items[o].ToString().Split(',')[0].Trim()) && (Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(0, 19)) <= Convert.ToDateTime(listBox9.Items[o].ToString().Split(',')[2].Trim())) && (Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(22, 19)) >= Convert.ToDateTime(listBox9.Items[o].ToString().Split(',')[3].Trim())))
                        {
                            if (num_progs[y].ToString() == listBox9.Items[o].ToString().Split(',')[1].Trim())
                            {
                                double maxtimesex = tsallpr.TotalSeconds * 0.7;
                                if (TimeSpan.Parse(listBox9.Items[o].ToString().Split(',')[4].Trim()) > TimeSpan.FromSeconds(maxtimesex)) { sum_otkl += TimeSpan.Parse(listBox9.Items[o].ToString().Split(',')[4].Trim()); }
                            }

                            // 

                        }

                        
                    }

                    day_prog += sum_otkl.ToString(); 


                }

                day_prog = day_prog.Remove(0, 1);
                day_prog = day_prog.Replace("O0 - 00:00:00; ", "");
                day_prog = day_prog.Replace("О0 - 00:00:00; ", "");

                double parametr_1 = TimeSpan.Parse("5:00:00").TotalSeconds;
                for (int y = 0; y < pause_ar.Count(); y++)
                {
                    if (pause_ar[y] < parametr_1)
                    {
                        pause_s += Convert.ToDouble(pause_ar[y]);
                    }

                    if (ostanov_ar[y] < parametr_1)
                    {
                        ostanov_s += Convert.ToDouble(ostanov_ar[y]);
                    }

                    //ostanov_s += Convert.ToDouble(ostanov_ar[y]);
                    //pause_s += Convert.ToDouble(pause_ar[y]);
                    mash_time += Convert.ToDouble(mash_ar[y]);
                    listBox3.Items.Add(pause_ar[y].ToString());
                    // MessageBox.Show(pause_ar[y].ToString()+"  -  " + ostanov_ar[y].ToString());
                   // MessageBox.Show(Convert.ToDouble(ostanov_ar[y]).ToString());
                }

                foreach (string num in num_progs)
                {
                    Regex regex = new Regex(@"O\d\d\d$");

                    MatchCollection matches = regex.Matches(num);
                    string num2 = num;
                    if (matches.Count > 0)
                    {
                        // MessageBox.Show("БЕДА");
                        num2 = num.Insert(1, "0");
                    }

                    List<TimeSpan> l_mash = new List<TimeSpan>();


                    int ko = 0;
                    double sr_mash_s = 0;
                    double sr_mash_s2 = 0;
                    for (int ii = 0; ii < dataGridView1.Rows.Count; ii++)
                    {
                        if ((dataGridView1["Программа", ii].Value.ToString() == num) && (dataGridView1["Счётчик", ii].Value.ToString() == "1"))
                        {
                            //MessageBox.Show("cyka");
                            if ((Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(0, 19)) <= Convert.ToDateTime(dataGridView1["Начало", ii].Value.ToString())) && (Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(22, 19)) >= Convert.ToDateTime(dataGridView1["Завершение", ii].Value.ToString())))
                            {
                                ko++;
                                l_mash.Add(TimeSpan.Parse(dataGridView1["Машинное время, ч.", ii].Value.ToString()));
                            }
                            else { }
                           

                        } 
                    }

                    foreach (TimeSpan l_mash_time in l_mash)
                    {
                        sr_mash_s += l_mash_time.TotalSeconds;
                    }

                    sr_mash_s2 = Math.Round(sr_mash_s / l_mash.Count,0);


                    num_progs2.Add(num2);
                    try
                    {
                        if (TimeSpan.TryParse(TimeSpan.FromSeconds(sr_mash_s2).ToString(), out TimeSpan result) == true)
                        {
                            programs += num2 + " - " + string.Format("{0:hh\\:mm\\:ss}", TimeSpan.FromSeconds(sr_mash_s2)) + " - " + ko.ToString() + "; ";
                        }
                    }
                    catch { }
                    
                       
                 
                }
                //try
              
                double osbl = Convert.ToDouble(ostanov_s);
                    double pausel = Convert.ToDouble(pause_s);
                    dataGridView4["Остановы, ч.", i].Value = Math.Round(osbl,2);
                    
                
                
                dataGridView4["Отклонения", i].Value = day_prog;

                dataGridView4["Пауза, ч.", i].Value = Convert.ToString(Math.Round(Convert.ToDouble(pausel), 2));
                    
                    double mash_time_up, mash_time_down;
                    mash_time_up = mash_time;
                    dataGridView4["Оператор", i].Value = comboBox2.Text;
                    dataGridView4["Номера программ", i].Value = programs;
                    dataGridView4["Машинное время, ч.", i].Value = Math.Round(Convert.ToDouble(mash_time_up) / 3600, 2);
                    mash_time_down = Convert.ToDouble(dataGridView4["Машинное время, ч.", i].Value) + Convert.ToDouble(dataGridView4["Остановы, ч.", i].Value) + Convert.ToDouble(dataGridView4["Пауза, ч.", i].Value);
                    dataGridView4["КПД, %", i].Value = Math.Round(Math.Round(Convert.ToDouble(mash_time_up) / 3600,2) / Math.Round(TimeSpan.Parse(dataGridView4["Длительность смены, ч.", i].Value.ToString()).TotalSeconds / 3600,2) * 100, 2);//(Convert.ToDouble(osbl + pausel + mash_time_up))) * 100, 2);

                    int group_stank = int.Parse(dataGridView4["Станок", i].Value.ToString()[1].ToString());

                    OleDbDataAdapter adapter = new OleDbDataAdapter();

                    string plastini=string.Empty;
                // OTPUSK KOM !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                //List<string> instr = new List<string>();
                //for (int t = 0; t < num_progs2.Count; t++)
                //{
                //    string query = "SELECT Пластина, Комментарий FROM [ПР-П] WHERE Программа = '" + num_progs2[t] + "'";// AND Группа = " + group_stank;
                //    OleDbCommand command = new OleDbCommand(query, myConnection);
                //    //MessageBox.Show(query);
                //    //string query = "SELECT Пластина FROM [ПР-П] WHERE Программа = '" + num_progs[t] + "' AND Группа = " + group_stank;
                //    //OleDbCommand command3 = new OleDbCommand(query);
                //    //adapter.SelectCommand = command;
                //    //DataSet dataSet = new DataSet();
                //    //adapter.Fill(dataSet);
                //    OleDbDataReader reader = command.ExecuteReader();
                //    while (reader.Read())
                //    {
                //       // MessageBox.Show(reader.GetString(0) + " - " + reader.GetString(1));
                //        try
                //        {
                //            //MessageBox.Show(reader.GetString(0)+" - "+ reader.GetString(1));
                //            if ((reader.GetString(0) == string.Empty) || (reader.GetString(0) ==""))
                //            {
                //                instr.Add(reader.GetString(1));
                //            }
                //            else
                //            {
                //                instr.Add(reader.GetString(0));
                //            }
                //        }
                //        catch { instr.Add(reader.GetString(1)); }

                //    }
                //   // reader.Close();
                //}

                //List<string> fin_instr = new List<string>();
                //fin_instr = instr.Distinct().ToList();

                //for (int tt = 0; tt < fin_instr.Count; tt++)
                //{
                //    plastini += fin_instr[tt]+"; ";
                //}

                //    dataGridView4["Инструмент", i].Value = plastini;

                // OTPUSK KOM !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                if (Convert.ToDouble(dataGridView4["КПД, %", i].Value) <= 50)
                    {
                        dataGridView4["КПД, %", i].Style.ForeColor = Color.Red;

                    }
                    else { dataGridView4["КПД, %", i].Style.ForeColor = Color.LimeGreen; }
                   
              
            }
            dataGridView4.Sort(dataGridView4.Columns["Период"], ListSortDirection.Ascending);
        }

        private void comboras(string combo, DataGridView data, System.Windows.Forms.TextBox outp1, System.Windows.Forms.TextBox outp2, System.Windows.Forms.TextBox outp3, System.Windows.Forms.TextBox outp4, System.Windows.Forms.Label min, System.Windows.Forms.Label max, System.Windows.Forms.Label on, System.Windows.Forms.Label off)
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
                try
                {
                    polez_s += Convert.ToInt32(dataGridView1["Числитель, с.", i].Value);
                    
                }
                catch { MessageBox.Show(polez_s + " | " + dataGridView1["Числитель, с.", i].Value.ToString()); }
                all += Convert.ToInt32(dataGridView1["Знаменатель, с.", i].Value);
            }
            itog = Math.Round(Convert.ToDouble(polez_s) / Convert.ToDouble((polez_s + all)) * 100, 2);
            outp4.Text = Convert.ToString(itog);
            ostanov_ar = new double[dataGridView1.RowCount];
            pause_ar = new double[dataGridView1.RowCount];
            outp1.Text = Convert.ToString(Math.Round(Convert.ToDouble(polez_s) / 3600, 2));
            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                ostanov_b_h = dataGridView1["Остановы, ч.", j].Value.ToString().Substring(0, 2);
                ostanov_b_m = dataGridView1["Остановы, ч.", j].Value.ToString().Substring(3, 2);
                ostanov_b_s = dataGridView1["Остановы, ч.", j].Value.ToString().Substring(6, 2);
                ostanov_s = Convert.ToInt32(ostanov_b_s);
                ostanov_s = ostanov_s + (Convert.ToInt32(ostanov_b_m) * 60);
                ostanov_s = ostanov_s + (Convert.ToInt32(ostanov_b_h) * 3600);
                if (dataGridView1["Пауза, ч.", j].Value.ToString().Length > 0)
                {
                    pause_b_h = dataGridView1["Пауза, ч.", j].Value.ToString().Substring(0, 2);
                    pause_b_m = dataGridView1["Пауза, ч.", j].Value.ToString().Substring(3, 2);
                    pause_b_s = dataGridView1["Пауза, ч.", j].Value.ToString().Substring(6, 2);
                    pause_s = 0;
                    if (!pause_b_s.Contains('-'))
                    {
                        pause_s = Convert.ToInt32(pause_b_s);
                        pause_s = pause_s + (Convert.ToInt32(pause_b_m) * 60);
                        pause_s = pause_s + (Convert.ToInt32(pause_b_h) * 3600);
                    }

                }
                ostanov_ar[j] = ostanov_s;
                pause_ar[j] = pause_s;
            }
            for (int y = 0; y < pause_ar.Count(); y++)
            {
                ostanov_s += Convert.ToInt32(ostanov_ar[y]);
                pause_s += Convert.ToInt32(pause_ar[y]);
            }

            outp2.Text = Convert.ToString(Math.Round(Convert.ToDouble(ostanov_s) / 3600, 2));
            outp3.Text = Convert.ToString(Math.Round(Convert.ToDouble(pause_s) / 3600, 2));
            stanok = new string[dataGridView1.Rows.Count];
            System.Data.DataTable dt1 = new System.Data.DataTable();
            DataRow dr1 = dt1.NewRow();
            data.Columns.Add("Станок", "Станок");

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
            catch { MessageBox.Show("Обнаружены некорректные значения даты."); }
            double on1 = 0, off1 = 0;

            try
            {
                for (int k = 0; k < dataGridView5.Rows.Count; k++)
                {
                    for (int y = 0; y < data.Rows.Count; y++)
                    {
                        if ((combo == dataGridView5[0, k].Value.ToString()) && (data[0, y].Value.ToString() == dataGridView5["Статус", k].Value.ToString()))
                        {
                            data[1, y].Value = dataGridView5[2, k].Value.ToString();
                            data[2, y].Value = dataGridView5[3, k].Value.ToString();
                            on1 += Double.Parse(data[1, y].Value.ToString(), CultureInfo.InvariantCulture) * 60;
                            off1 += Double.Parse(data[2, y].Value.ToString(), CultureInfo.InvariantCulture) * 60;
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
                        string f = "";
                        for (int h = 0; h < data.Rows.Count; h++)
                        {

                            f += dataGridView5.Columns["Станок"].HeaderText.ToString() + " LIKE '%" + data[0, h].Value.ToString() + "%' AND " + dataGridView5.Columns["Статус"].HeaderText.ToString() + " LIKE 'Станок выключен' AND Начало > '" + Convert.ToDateTime(minValue_1) + "' and Конец < '" + Convert.ToDateTime(maxValue_1) + "' OR "; 

                        }
                        int er = f.LastIndexOf("OR");
                        f = f.Substring(0, f.Length - 3);
                        f += "AND " + dataGridView5.Columns["Статус"].HeaderText.ToString() + " LIKE 'Станок выключен'";
                        bs_d5.Filter = f.ToString();
                        dataGridView5.DataSource = bs_d5;
                        dataGridView5.Sort(dataGridView5.Columns["Станок"], ListSortDirection.Ascending);
                    }
                    catch { }
                }
            }
            catch { }
            

        }




        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void comboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            dataGridView2.Columns.Clear();
            comboras(comboBox1.Text, dataGridView2, textBox1, textBox2, textBox3, textBox6, label17, label18, label22, label23);
            
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            listBox5.Items.Clear();
            listBox6.Items.Clear();
            listBox7.Items.Clear();
            listBox8.Items.Clear();
            listBox9.Items.Clear();

            repoday(comboBox2.Text);
            double seconds = 0, sec = 0;
            double col = 0;
            List<DateTime> l_day = new List<DateTime>();
            List<DateTime> l_day_beg = new List<DateTime>();
            List<DateTime> l_day_end = new List<DateTime>();
            List<DateTime> cringe1 = new List<DateTime>();
            List<DateTime> cringe2 = new List<DateTime>();
            List<DateTime> unikend_ar = new List<DateTime>();
            List<DateTime> unikbeg_ar = new List<DateTime>();
            List<DateTime> smena1beg_ar = new List<DateTime>();
            List<DateTime> smena1end_ar = new List<DateTime>();
            List<DateTime> smena2beg_ar = new List<DateTime>();
            List<DateTime> smena2end_ar = new List<DateTime>();
            List<DateTime> smena3beg_ar = new List<DateTime>();
            List<DateTime> smena3end_ar = new List<DateTime>();

            List<string> ishod = new List<string>();
            bool haveis = false;

            dataGridView6.Rows.Clear();
            listBox5.Items.Clear();

            List<int> l_kolvo = new List<int>();

            // СОЗДАЁМ МАССИВ ДНЕЙ
            try
            {
                l_day.Add(Convert.ToDateTime(dataGridView4["День", 0].Value.ToString()).Date);
                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    bool day_is = false;
                    l_day_beg.Add(Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(0, 19)));
                    l_day_end.Add(Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(22, 19)));
                    for (int y = 0; y < l_day.Count; y++)
                    {
                        if (l_day[y] == Convert.ToDateTime(dataGridView4["День", i].Value.ToString()).Date)
                        {
                            day_is = true;
                        }
                    }
                    if (day_is == false)
                    {
                        l_day.Add(Convert.ToDateTime(dataGridView4["День", i].Value.ToString()));
                    }
                }
                // МАССИВ ДНЕЙ ЗАПОЛНЕН
            }
            catch { }


            // РАБОТА
            for (int day_ind = 0; day_ind < l_day.Count; day_ind++)
            {
                //int cnt = 0;
                List<DateTime> iday_ar = new List<DateTime>();
                List<DateTime> idayend_ar = new List<DateTime>();


                // ЗАПОЛНЯЕМ МАССИВ НАЧАЛЬНЫХ ДАТ
                for (int beg_ind = 0; beg_ind < l_day_beg.Count; beg_ind++)
                {
                    if ((l_day[day_ind].Day == l_day_beg[beg_ind].Day) && (l_day[day_ind].Month == l_day_beg[beg_ind].Month) && (l_day[day_ind].Year == l_day_beg[beg_ind].Year))
                    {
                        //cnt++;
                        iday_ar.Add(l_day_beg[beg_ind]);
                        idayend_ar.Add(l_day_end[beg_ind]);
                    }
                    //
                }
                // МАССИВ ЗАПОЛНЕН

                // ОПРЕДЕЛЕНИЕ ЗНАЧЕНИЙ НАЧАЛА СМЕН
                DateTime beg_min = iday_ar[0];
                DateTime fuckingday = iday_ar[0].Date;
                DateTime beg_max = iday_ar[0];
                for (int y = 0; y < iday_ar.Count; y++)
                {
                    DateTime base_date = iday_ar[y].Date;
                    //(beg_min > iday_ar[y]) && 
                    if ((iday_ar[y] - base_date >= TimeSpan.Parse("0:00:00")) && (iday_ar[y] - base_date <= TimeSpan.Parse("14:30:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:00:00")))
                    {
                        smena1beg_ar.Add(iday_ar[y]);
                        smena1end_ar.Add(idayend_ar[y]);

                    }

                    if ((iday_ar[y] - base_date >= TimeSpan.Parse("14:30:00")) && (iday_ar[y] - base_date <= TimeSpan.Parse("23:30:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:50:00")))
                    {
                        smena2beg_ar.Add(iday_ar[y]);
                        smena2end_ar.Add(idayend_ar[y]);
                    }

                    //if ((iday_ar[y] - base_date >= TimeSpan.Parse("23:30:00")) && (iday_ar[y] - base_date <= TimeSpan.Parse("1.06:30:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:00:00")))
                    //{
                    //    smena3beg_ar.Add(iday_ar[y]);
                    //    smena3end_ar.Add(idayend_ar[y]);
                    //}
                }

                for (int y = 0; y < iday_ar.Count; y++)
                {
                    DateTime base_date = iday_ar[y].Date;
                    if ((iday_ar[y] - base_date >= TimeSpan.Parse("7:00:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("15:30:00")))// && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:00:00")))
                    {
                        smena1beg_ar.Add(iday_ar[y]);
                        smena1end_ar.Add(idayend_ar[y]);
                    }

                    if ((iday_ar[y] - base_date >= TimeSpan.Parse("15:30:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("23:59:59")))// && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:50:00")))
                    {
                        smena2beg_ar.Add(iday_ar[y]);
                        smena2end_ar.Add(idayend_ar[y]);
                    }

                    if ((iday_ar[y] - base_date >= TimeSpan.Parse("0:00:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("07:00:00")))// && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:00:00")))
                    {
                        smena3beg_ar.Add(iday_ar[y]);
                        smena3end_ar.Add(idayend_ar[y]);
                    }
                }

                // Допиливаем Смены

                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {

                    DateTime dt_day_b = Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Split('-')[0]);
                    DateTime dt_day_e = Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Split('-')[1]);

                    formsmen(dt_day_e, dt_day_b, smena1beg_ar, smena2beg_ar, smena3beg_ar, smena1end_ar, smena2end_ar, smena3end_ar);

                    if (dt_day_b - dt_day_b.Date > TimeSpan.Parse("0:00:00") && dt_day_b - dt_day_b.Date < TimeSpan.Parse("7:00:00") && dt_day_e - dt_day_b.Date > TimeSpan.Parse("7:00:00"))
                    {

                        DateTime new_end = dt_day_b.Date + TimeSpan.Parse("7:00:00");
                        if (!smena3beg_ar.Contains(dt_day_b))
                        {
                            if (dt_day_e - dt_day_b.Date < TimeSpan.Parse("15:30:00"))
                            {
                                smena3beg_ar.Add(dt_day_b);
                                smena3end_ar.Add(new_end);
                            }
                            if (dt_day_e - dt_day_b.Date < TimeSpan.Parse("15:30:00"))
                            {
                                smena2beg_ar.Add(new_end);
                                smena2end_ar.Add(dt_day_e);
                            }
                            if (dt_day_e - dt_day_b.Date < TimeSpan.Parse("1.00:00:00"))
                            {
                                smena1beg_ar.Add(new_end);
                                smena1end_ar.Add(Convert.ToDateTime("15:30:00"));
                                smena2beg_ar.Add(new_end);
                                smena2end_ar.Add(Convert.ToDateTime("15:30:00"));
                            }

                        }


                        if (dt_day_e - dt_day_b.Date > TimeSpan.Parse("7:00:00") && dt_day_e - dt_day_b.Date < TimeSpan.Parse("15:30:00"))
                        {
                            DateTime new_end_1 = dt_day_b.Date + TimeSpan.Parse("15:30:00");
                            DateTime new_beg_1 = dt_day_b.Date + TimeSpan.Parse("07:00:00");
                            smena1beg_ar.Add(new_beg_1);
                            smena1end_ar.Add(new_end_1);
                        }

                        if (dt_day_e - dt_day_b.Date > TimeSpan.Parse("15:30:00") && dt_day_e - dt_day_b.Date < TimeSpan.Parse("23:59:59"))
                        {
                            DateTime new_end_1 = dt_day_b.Date + TimeSpan.Parse("23:59:59");
                            DateTime new_beg_1 = dt_day_b.Date + TimeSpan.Parse("15:30:00");
                            smena2beg_ar.Add(new_beg_1);
                            smena2end_ar.Add(new_end_1);
                        }
                    }
                    // }

                    if (dt_day_b - dt_day_b.Date > TimeSpan.Parse("15:30:00") && dt_day_b - dt_day_b.Date < TimeSpan.Parse("23:59:59") && dt_day_e - dt_day_b.Date > TimeSpan.Parse("23:59:59"))
                    {
                        DateTime new_end = dt_day_b.Date + TimeSpan.Parse("23:59:59");
                        smena2beg_ar.Add(dt_day_b);
                        smena2end_ar.Add(new_end);

                        if (dt_day_e - dt_day_b.Date > TimeSpan.Parse("23:59:59") && dt_day_e - dt_day_b.Date < TimeSpan.Parse("1.07:00:00"))
                        {
                            DateTime new_end_1 = dt_day_b.Date + TimeSpan.Parse("1.07:00:00");
                            DateTime new_beg_1 = dt_day_b.Date + TimeSpan.Parse("1.00:00:00");
                            smena3beg_ar.Add(new_beg_1);
                            smena3end_ar.Add(new_end_1);
                        }

                        if (dt_day_e - dt_day_b.Date > TimeSpan.Parse("1.07:00:00") && dt_day_e - dt_day_b.Date < TimeSpan.Parse("1.15:30:00"))
                        {
                            DateTime new_end_1 = dt_day_b.Date + TimeSpan.Parse("1.15:30:00");
                            DateTime new_beg_1 = dt_day_b.Date + TimeSpan.Parse("1.07:00:00");
                            smena1beg_ar.Add(new_beg_1);
                            smena1end_ar.Add(new_end_1);
                        }
                    }

                    if (dt_day_b - dt_day_b.Date > TimeSpan.Parse("7:00:00") && dt_day_b - dt_day_b.Date < TimeSpan.Parse("15:30:00") && dt_day_e - dt_day_b.Date > TimeSpan.Parse("15:30:00"))
                    {
                        DateTime new_end = dt_day_b.Date + TimeSpan.Parse("15:30:00");
                        smena1beg_ar.Add(dt_day_b);
                        smena1end_ar.Add(new_end);

                        if (dt_day_e - dt_day_b.Date > TimeSpan.Parse("15:30:00") && dt_day_e - dt_day_b.Date < TimeSpan.Parse("23:59:59"))
                        {
                            DateTime new_end_1 = dt_day_b.Date + TimeSpan.Parse("23:59:59");
                            DateTime new_beg_1 = dt_day_b.Date + TimeSpan.Parse("15:30:00");
                            smena2beg_ar.Add(new_beg_1);
                            smena2end_ar.Add(new_end_1);
                        }

                        if (dt_day_e - dt_day_b.Date > TimeSpan.Parse("23:59:59") && dt_day_e - dt_day_b.Date < TimeSpan.Parse("1.07:00:00"))
                        {
                            DateTime new_end_1 = dt_day_b.Date + TimeSpan.Parse("1.07:00:00");
                            DateTime new_beg_1 = dt_day_b.Date + TimeSpan.Parse("1.00:00:00");
                            smena3beg_ar.Add(new_beg_1);
                            smena3end_ar.Add(new_end_1);
                        }
                    }


                }
            }

            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {

                DateTime dt_day_b = Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Split('-')[0]);
                DateTime dt_day_e = Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Split('-')[1]);

                formsmen(dt_day_e, dt_day_b, smena1beg_ar, smena2beg_ar, smena3beg_ar, smena1end_ar, smena2end_ar, smena3end_ar);
            }

            for (int i = 0; i < smena1beg_ar.Count; i++)
            {
                for (int j = 0; j < smena2beg_ar.Count; j++)
                {
                    if ((smena1beg_ar[i].Date == smena2beg_ar[j].Date) && (smena2beg_ar[j] - smena1beg_ar[i] < TimeSpan.Parse("3:00:00")) && (smena1end_ar[i] > smena2end_ar[j]))
                    {
                        smena2beg_ar[j] = smena1beg_ar[i];
                        smena1beg_ar.Remove(smena1beg_ar[i]);
                        smena1end_ar.Remove(smena1end_ar[i]);
                    }
                }
            }

            for (int u = 0; u < smena3beg_ar.Count; u++)
            {
                listBox6.Items.Add(smena3beg_ar[u].ToString() + " - " + smena3end_ar[u].ToString());
            }


            //List<string> fuck = new List<string>();
            List<DateTime> list_begday = new List<DateTime>();
            List<DateTime> list_endday = new List<DateTime>();
            List<DateTime> list_begday2 = new List<DateTime>();
            List<DateTime> list_endday2 = new List<DateTime>();
            List<DateTime> list_begday3 = new List<DateTime>();
            List<DateTime> list_endday3 = new List<DateTime>();
            List<DateTime> c1 = new List<DateTime>();
            List<DateTime> c2 = new List<DateTime>();

            List<DateTime> c1_2 = new List<DateTime>();
            List<DateTime> c2_2 = new List<DateTime>();

            List<DateTime> cccb1 = new List<DateTime>();
            List<DateTime> ccce1 = new List<DateTime>();
            List<DateTime> cccb2 = new List<DateTime>();
            List<DateTime> ccce2 = new List<DateTime>();


            List<DateTime> indeces_b = new List<DateTime>();
            List<DateTime> indeces_e = new List<DateTime>();

            List<DateTime> indeces_b2 = new List<DateTime>();
            List<DateTime> indeces_e2 = new List<DateTime>();



            List<int> cccol = new List<int>();

            List<DateTime> b1 = smena1beg_ar.ToList();
            List<DateTime> b2 = smena2beg_ar.ToList();
            List<DateTime> e1 = smena1end_ar.ToList();
            List<DateTime> e2 = smena2end_ar.ToList();


            for (int day = 0; day < l_day.Count; day++)
            {
                DateTime beg_max = DateTime.MinValue;
                DateTime beg_min = DateTime.MaxValue;
                DateTime beg_max2 = DateTime.MinValue;
                DateTime beg_min2 = DateTime.MaxValue;
                DateTime beg_max3 = DateTime.MinValue;
                DateTime beg_min3 = DateTime.MaxValue;

                // ОТЧАЯНИЕ
                DateTime A2beg_max = DateTime.MaxValue;
                DateTime A2beg_min = DateTime.MinValue;
                DateTime A2beg_max2 = DateTime.MaxValue;
                DateTime A2beg_min2 = DateTime.MinValue;
                DateTime A2beg_max3 = DateTime.MaxValue;
                DateTime A2beg_min3 = DateTime.MinValue;


                DateTime A2beg_max_2 = DateTime.MaxValue;
                DateTime A2beg_min_2 = DateTime.MinValue;
                DateTime A2beg_max2_2 = DateTime.MaxValue;
                DateTime A2beg_min2_2 = DateTime.MinValue;
                DateTime A2beg_max3_2 = DateTime.MaxValue;
                DateTime A2beg_min3_2 = DateTime.MinValue;

                int kolvo_st = 0;
                int day_ogr = 0;

                for (int i = 0; i < smena1beg_ar.Count; i++)
                {
                    if (l_day[day].Date == smena1beg_ar[i].Date)
                    {
                        kolvo_st++;
                        day_ogr = i;

                        if (beg_min > smena1beg_ar[i])
                        {
                            beg_min = smena1beg_ar[i];
                        }

                        if (beg_max < smena1end_ar[i])
                        {
                            beg_max = smena1end_ar[i];
                        }
                    }
                }

                for (int i = 0; i < smena2beg_ar.Count; i++)
                {
                    if (l_day[day].Date == smena2beg_ar[i].Date)
                    {
                        kolvo_st++;
                        if (beg_min2 > smena2beg_ar[i])
                        {
                            beg_min2 = smena2beg_ar[i];
                        }

                        if (beg_max2 < smena2end_ar[i])
                        {
                            beg_max2 = smena2end_ar[i];
                        }
                    }
                }


                for (int i = 0; i < b1.Count; i++)
                {

                    if (l_day[day].Date == b1[i].Date)
                    {
                        if (A2beg_min < b1[i])
                        {
                            if (indeces_b.Contains(b1[i]) == false)
                            {
                                A2beg_min = b1[i];
                                indeces_b.Add(b1[i]);
                            }
                            else
                            {
                                A2beg_min = b1[i];
                            }
                        }
                    }


                    if (l_day[day].Date == b1[i].Date)
                    {

                        if (A2beg_max > e1[i])
                        {
                            if (indeces_b.Contains(e1[i]) == false)
                            {
                                A2beg_max = e1[i];
                                indeces_e.Add(e1[i]);
                            }
                            else
                            {
                                A2beg_max = e1[i];
                            }
                        }
                    }

                }
                for (int i = 0; i < b2.Count; i++)
                {

                    if (l_day[day].Date == b2[i].Date)
                    {
                        if (A2beg_min2 < b2[i])
                        {
                            if (indeces_b2.Contains(b2[i]) == false)
                            {
                                A2beg_min2 = b2[i];
                                indeces_b2.Add(b2[i]);
                            }
                            else
                            {
                                A2beg_min2 = b2[i];
                            }
                        }
                    }


                    if (l_day[day].Date == b2[i].Date)
                    {
                        if (A2beg_max2 > e2[i])
                        {
                            if (indeces_b2.Contains(e2[i]) == false)
                            {
                                A2beg_max2 = e2[i];
                                indeces_e2.Add(e2[i]);
                            }
                            else
                            {
                                A2beg_max2 = e2[i];
                            }
                        }

                    }
                }


                if (beg_min.Date == l_day[day].Date)
                {

                    list_begday.Add(beg_min);
                    list_endday.Add(beg_max);
                    c1.Add(A2beg_max);
                    c2.Add(A2beg_min);
                    cccol.Add(kolvo_st);
                }

                if (beg_min2.Date == l_day[day].Date)
                {
                    list_begday2.Add(beg_min2);
                    list_endday2.Add(beg_max2);
                    c1.Add(A2beg_max2);
                    c2.Add(A2beg_min2);
                    cccol.Add(kolvo_st);
                }



            }
            var alltime_beg = c2.OrderBy(s => s);
            var alltime_end = c1.OrderBy(s => s);

            List<DateTime> s1_b = new List<DateTime>();
            List<DateTime> s1_e = new List<DateTime>();
            List<DateTime> s1_ef1 = new List<DateTime>();
            List<DateTime> s2_b = new List<DateTime>();
            List<DateTime> s2_e = new List<DateTime>();
            List<DateTime> s2_ef1 = new List<DateTime>();
            List<int> s1_count = new List<int>();
            List<int> s2_count = new List<int>();




            //for (int cstanki=6; cstanki>0; cstanki--)
            //{
            for (int day = 0; day < l_day.Count; day++)
            {
                List<DateTime> stanok1_b = new List<DateTime>();
                List<DateTime> stanok1_e = new List<DateTime>();
                List<DateTime> stanok2_b = new List<DateTime>();
                List<DateTime> stanok2_e = new List<DateTime>();
                List<DateTime> stanok1_ef1 = new List<DateTime>();
                List<int> stc_e = new List<int>();
                List<int> stc_e_2 = new List<int>();

                DateTime st1_beg = DateTime.MinValue;
                DateTime st1_end = DateTime.MaxValue;
                DateTime st2_beg = DateTime.MinValue;
                DateTime st2_end = DateTime.MaxValue;

                int kolvo_st = 0;
                int kolvo_st_2 = 0;

                for (int row = 0; row < smena1beg_ar.Count; row++)
                {
                    DateTime dg_row_b = smena1beg_ar[row];
                    DateTime dg_row_e = smena1end_ar[row];
                    if (l_day[day].Date == dg_row_b.Date)
                    {

                        kolvo_st++;

                        stanok1_b.Add(dg_row_b);
                        stanok1_e.Add(dg_row_e);
                        stc_e.Add(kolvo_st);
                    }

                }
                for (int row = 0; row < smena2beg_ar.Count; row++)
                {

                    DateTime dg_row_b2 = smena2beg_ar[row];
                    DateTime dg_row_e2 = smena2end_ar[row];
  
                    if (l_day[day].Date == dg_row_b2.Date)
                    {

                        kolvo_st_2++;

                        stanok2_b.Add(dg_row_b2);
                        stanok2_e.Add(dg_row_e2);
                        stc_e_2.Add(kolvo_st_2);
                    }

                }

                var sort_st_b = stanok1_b.OrderBy(s => s);
                var sort_st_e = stanok1_e.OrderByDescending(s => s);
                var sort_st_b2 = stanok2_b.OrderBy(s => s);
                var sort_st_e2 = stanok2_e.OrderByDescending(s => s);



                s1_b.AddRange(sort_st_b);
                s1_e.AddRange(sort_st_e);
                s1_count.AddRange(stc_e);
                s2_b.AddRange(sort_st_b2);
                s2_e.AddRange(sort_st_e2);
                s2_count.AddRange(stc_e_2);

            }
            //}
            var albegtime = s1_b.Concat(s2_b).OrderBy(s => s);
            var alendtime = s1_e.Concat(s2_e).OrderBy(s => s);

            List<DateTime> ABT = albegtime.ToList();
            List<DateTime> AET = alendtime.ToList();
            List<int> AC = new List<int>();

            int kol = 0;

            List<string> part_1 = new List<string>();
            try
            {
                for (int row = 0; row < s1_b.Count - 1; row++)
                {
                    if (s1_count[row] == 1 && s1_b[row].Date == s1_b[row + 1].Date)
                    {
                        part_1.Add(s1_b[row].ToString() + " - " + s1_b[row + 1].ToString() + " - " + s1_count[row].ToString());
                    }
                    else if (s1_b[row].Date != s1_b[row + 1].Date)// && row < s1_b.Count-2)
                    {
                        part_1.Add(s1_b[row].ToString() + " - " + s1_e[row].ToString() + " - " + s1_count[row].ToString());
                    }

                }
                part_1.Add(s1_b[s1_b.Count - 1].ToString() + " - " + s1_e[s1_b.Count - 1].ToString() + " - " + s1_count[s1_b.Count - 1].ToString());

                for (int row = 0; row < s2_b.Count - 1; row++)
                {
                    if (s2_count[row] == 1 && s2_b[row + 1].Date == s2_b[row].Date)
                    {
                        part_1.Add(s2_b[row].ToString() + " - " + s2_b[row + 1].ToString() + " - " + s2_count[row].ToString());
                    }
                    else if (s2_b[row + 1].Date != s2_b[row].Date)
                    {
                        part_1.Add(s2_b[row].ToString() + " - " + s2_e[row].ToString() + " - " + s2_count[row].ToString());
                        //part_1.Add(s2_b[s2_b.Count].ToString() + " - " + s2_e[s2_b.Count].ToString() + " - " + s2_count[s2_b.Count].ToString());
                    }
                }
                part_1.Add(s2_b[s2_b.Count - 1].ToString() + " - " + s2_e[s2_b.Count - 1].ToString() + " - " + s2_count[s2_b.Count - 1].ToString());
            }
            catch { }

            part_1.Sort();

            foreach (string lines in listBox5.Items)
            {
                ishod.Add(lines);
            }

            List<string> list = new List<string>();

            System.Data.DataTable prostoi = new System.Data.DataTable();

            prostoi.Columns.Add("День");
            prostoi.Columns.Add("Станок");
            prostoi.Columns.Add("Программа");
            prostoi.Columns.Add("Простой");
            prostoi.Columns.Add("Начало");
            prostoi.Columns.Add("Конец");
            prostoi.Columns["Начало"].DataType = Type.GetType("System.DateTime");
            prostoi.Columns["Конец"].DataType = Type.GetType("System.DateTime");

            List<string> num_prog = new List<string>();
            List<string> L_stanok = new List<string>();
            List<string> prost = new List<string>();

            for (int day = 0; day < l_day.Count; day++)
            {
                //MessageBox.Show(Convert.ToDateTime(l_day[day].Date).Date.ToString());
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataRow dr_prostoi = prostoi.NewRow();
                    if (dataGridView1["Начало", i].Value.ToString().Substring(0, 10) == l_day[day].ToString().Substring(0, 10))
                    {
                        dr_prostoi[0] = dataGridView1["Начало", i].Value.ToString().Substring(0, 10);
                        dr_prostoi[1] = dataGridView1["Оборудование", i].Value.ToString();
                        int value;
                        int.TryParse(string.Join("", dataGridView1["Программа", i].Value.ToString().Where(c => char.IsDigit(c))), out value);
                        dr_prostoi[2] = value;
                        dr_prostoi[3] = dataGridView1["Пауза, ч.", i].Value.ToString();
                        dr_prostoi[4] = Convert.ToDateTime(dataGridView1["Начало", i].Value.ToString());
                        dr_prostoi[5] = Convert.ToDateTime(dataGridView1["Завершение", i].Value.ToString());
                        prostoi.Rows.Add(dr_prostoi);
                    }
                }

            }

            DataView dv = prostoi.DefaultView;
            dv.Sort = "Станок, Начало asc";
            DataTable sorted_prostoi = dv.ToTable();
            dataGridView7.DataSource = sorted_prostoi;
            dataGridView7.Columns["Начало"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm:ss";
            dataGridView7.Columns["Конец"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm:ss";
            List<string> program = new List<string>();
            List<string> st = new List<string>();
            List<string> day_l = new List<string>();
            List<string> last = new List<string>();
            string last_row = "";

            for (int i = 0; i < dataGridView7.Rows.Count - 1; i++)
            {
                if (!program.Contains(dataGridView7[2, i].Value.ToString()))
                {
                    program.Add(dataGridView7[2, i].Value.ToString());
                }

                if (!st.Contains(dataGridView7[1, i].Value.ToString()))
                {
                    st.Add(dataGridView7[1, i].Value.ToString());
                }

                if (!day_l.Contains(dataGridView7[0, i].Value.ToString()))
                {
                    day_l.Add(dataGridView7[0, i].Value.ToString());
                }
            }

            for (int d = 0; d < day_l.Count; d++)
            {
                for (int ist = 0; ist < st.Count; ist++)
                {
                    List<string> nm_prog_done = new List<string>();

                    for (int dr = 0; dr < dataGridView7.Rows.Count - 1; dr++)
                    {
                        if (st[ist] == dataGridView7[1, dr].Value.ToString())
                        {
                            nm_prog_done.Add(dataGridView7[2, dr].Value.ToString());
                            if ((day_l[d] == dataGridView7[0, dr].Value.ToString()) || (!nm_prog_done.Contains(dataGridView7[2, dr].Value.ToString())))
                            {
                                nm_prog_done.Add(dataGridView7[2, dr].Value.ToString());
                                last_row = dataGridView7[1, dr].Value.ToString() + "|" + dataGridView7[4, dr].Value.ToString() + "|" + dataGridView7[5, dr].Value.ToString() + "|" + dataGridView7[2, dr].Value.ToString();
                                last.Add(dataGridView7[1, dr].Value.ToString() + "|" + dataGridView7[4, dr].Value.ToString() + "|" + dataGridView7[5, dr].Value.ToString() + "|" + dataGridView7[2, dr].Value.ToString());
                            }
                        }
                    }
                }
            }


            for (int i = 1; i < last.Count; i++)
            {
                try
                {
                    if ((last[i - 1].ToString().Split('|')[0] == last[i].ToString().Split('|')[0]))
                    {
                        if ((last[i - 1].ToString().Split('|')[3] != last[i].ToString().Split('|')[3]) && ((last[i - 1].ToString().Split('|')[3] != "0") && (last[i].ToString().Split('|')[0] == last[i - 1].ToString().Split('|')[0])))
                        {
                            listBox7.Items.Add(last[i - 1]); //+ " - " + prostoi_ts.ToString());
                        }
                        else if ((last[i - 1].ToString().Split('|')[3] != last[i].ToString().Split('|')[3]) && ((last[i - 1].ToString().Split('|')[3] != "0")) && (Convert.ToDateTime(last[i - 1].ToString().Split('|')[1]).Date <= Convert.ToDateTime(last[i - 1].ToString().Split('|')[1]).Date))
                        { listBox7.Items.Add(last[i - 1]); }
                    }

                }
                catch { }
                listBox8.Items.Add(last[i - 1] + "    -     " + last[i]);
            }

            dv.Sort = "Станок, Начало asc";
            DataTable sorted_prostoi1 = dv.ToTable();
            dataGridView8.DataSource = sorted_prostoi1;

            dataGridView8.Columns["Начало"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm:ss";
            dataGridView8.Columns["Конец"].DefaultCellStyle.Format = "dd.MM.yyyy HH:mm:ss";

            List<string> L_prst = new List<string>();
            List<int> rows_bord = new List<int>();

            for (int r = 0; r < listBox7.Items.Count; r++)
            {
                TimeSpan prst = TimeSpan.Zero;
                for (int i = 0; i < sorted_prostoi1.Rows.Count - 1; i++)
                {
                    //MessageBox.Show(sorted_prostoi1[0,i]);
                    if (sorted_prostoi1.Rows[i][1].ToString() == sorted_prostoi1.Rows[i + 1][1].ToString())
                    {
                        if ((listBox7.Items[r].ToString().Split('|')[0] == sorted_prostoi1.Rows[i][1].ToString()) && (listBox7.Items[r].ToString().Split('|')[1] == sorted_prostoi1.Rows[i][4].ToString()) && (listBox7.Items[r].ToString().Split('|')[2] == sorted_prostoi1.Rows[i][5].ToString()) && (listBox7.Items[r].ToString().Split('|')[3] == sorted_prostoi1.Rows[i][2].ToString()))
                        {
                            if ((sorted_prostoi1.Rows[i + 1][0].ToString() == sorted_prostoi1.Rows[i][0].ToString()) || (Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString()) - Convert.ToDateTime(sorted_prostoi1.Rows[i][5].ToString()) <= TimeSpan.Parse("05:00:00")))
                            {
                                try
                                {
                                    if (((Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString())) - (Convert.ToDateTime(sorted_prostoi1.Rows[i][5].ToString())) >= TimeSpan.Parse("10:00:00")) && Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString()).TimeOfDay >= TimeSpan.Parse("15:30:00"))
                                    {
                                        prst = Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString()) - Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][5].ToString().Substring(0, 10) + " 15:30:00");
                                        dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                                    }
                                    else if ((sorted_prostoi1.Rows[i + 1][1].ToString() == sorted_prostoi1.Rows[i][1].ToString()))
                                    {
                                        prst = Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString()) - Convert.ToDateTime(sorted_prostoi1.Rows[i][5].ToString());
                                        dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
                                        rows_bord.Add(i);
                                    }
                                }
                                catch { }

                            }
                            else if (Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString()) - Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][5].ToString()).Date >= TimeSpan.Parse("15:30:00"))
                            {
                                prst = Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString()) - Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][5].ToString().Substring(0, 10) + " 15:30:00");
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.Green;
                            }
                            else
                            {
                                prst = Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][4].ToString()) - Convert.ToDateTime(sorted_prostoi1.Rows[i + 1][5].ToString().Substring(0, 10) + " 07:00:00");
                                dataGridView8.Rows[i].DefaultCellStyle.BackColor = Color.Blue;
                            }


                            L_prst.Add(sorted_prostoi1.Rows[i + 1][2].ToString() + " - " + prst.ToString());
                        }
                    }

                    //MessageBox.Show(sorted_prostoi1.Rows[i][1].ToString());
                }
            }

            

            listBox8.Items.Clear();

            if (listBox7.Items.Count > 0)
            {
                List<string> uniq_prost = new List<string>();
                List<string> l_prost = new List<string>();

                for (int i = 0; i < listBox7.Items.Count; i++)
                {
                    try
                    {
                        listBox7.Items[i] += "-" + L_prst[i];
                        //try
                        //{

                        if ((TimeSpan.Parse(Convert.ToString(listBox7.Items[i].ToString()).Split('|')[3].Split('-')[2].Trim()) <= TimeSpan.Parse("00:00:00")))
                        {
                            listBox7.Items.Remove(listBox7.Items[i]);
                        }
                    }
                    catch (Exception)
                    {

                        listBox7.Items[i] = listBox7.Items[i].ToString().Replace(" - -", " - ");
                        //MessageBox.Show(listBox7.Items[i].ToString()); 
                    }

                    // }catch (Exception) { listBox7.Items.RemoveAt(i); } // Иногда в Лист добавляются лишние значения. Лучше исправить наполнение listbox7

                    try
                    {

                        if (listBox7.Items[i].ToString().Split('|')[3].Split('-').Count() < 3)
                        {
                        
                            listBox7.Items.Remove(listBox7.Items[i]);
                        
                        }
                    }
                    catch { // ОБРАБОТАТЬ ИСКЛЮЧЕНИЕ 11.02.2025
                            }


                    // ПОДСЧЕТ ИТОГОВ ПРОСТОЕВ
                    try
                    {
                        if (!uniq_prost.Contains(listBox7.Items[i].ToString().Split('|')[0] + " | " + listBox7.Items[i].ToString().Split('|')[1].Split(' ')[0] + " | " + listBox7.Items[i].ToString().Split('|')[3].Remove(listBox7.Items[i].ToString().Split('|')[3].LastIndexOf('-'))))
                        {
                            uniq_prost.Add(listBox7.Items[i].ToString().Split('|')[0] + " | " + listBox7.Items[i].ToString().Split('|')[1].Split(' ')[0] + " | " + listBox7.Items[i].ToString().Split('|')[3].Remove(listBox7.Items[i].ToString().Split('|')[3].LastIndexOf('-')));
                            //listBox8.Items.Add(listBox7.Items[i].ToString().Split('|')[0] + " | " + listBox7.Items[i].ToString().Split('|')[1].Split(' ')[0] + " | " + listBox7.Items[i].ToString().Split('|')[3].Remove(listBox7.Items[i].ToString().Split('|')[3].LastIndexOf('-')));
                        }
                    }
                    catch { }
                   

                }

                for (int n = 0; n < uniq_prost.Count; n++)
                {
                    TimeSpan prostoi_itog = TimeSpan.Zero;
                    for (int i = 0; i < listBox7.Items.Count; i++)
                    {
                        
                        
                        if (i<listBox7.Items.Count - 1)
                        {
                            string[] nv = listBox7.Items[i].ToString().Split('|');
                            string[] un = uniq_prost[n].ToString().Split('|');
                            if ((un[0] == nv[0]) && (un[1] == nv[1].Split(' ')[0]) && (un[2] == nv[3].Remove(nv[3].LastIndexOf('-'))))
                            {
                                prostoi_itog += TimeSpan.Parse(nv[3].Trim().Substring(nv[3].LastIndexOf('-'),8));
                                
                            }
                        }
                    }
                    listBox8.Items.Add(uniq_prost[n]+" = "+ prostoi_itog.ToString());
                }

            }

            

        }
        private void formsmen(DateTime dk, DateTime dn, List<DateTime> sm1, List<DateTime> sm2, List<DateTime> sm3, List<DateTime> sm1_e, List<DateTime> sm2_e, List<DateTime> sm3_e)
        {
            if (dn < dn.Date + TimeSpan.Parse("7:00:00") && dn > dn.Date)
            {
                if (dk > dn.Date + TimeSpan.Parse("7:00:00") && dk < dn.Date + TimeSpan.Parse("15:30:00"))
                {
                    if (!sm3.Contains(dn) && !sm3_e.Contains(dk))
                    {
                        sm1.Add(dn.Date + TimeSpan.Parse("7:00:00"));
                        sm1_e.Add(dk);
                        sm3.Add(dn);
                        sm3_e.Add(dn.Date + TimeSpan.Parse("7:00:00"));
                    }

                }
                else if (dk > dn.Date + TimeSpan.Parse("15:30:00") && dk < dn.Date + TimeSpan.Parse("23:59:59"))
                {
                    if (!sm3.Contains(dn) && !sm3_e.Contains(dk))
                    {
                        sm2.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                        sm2_e.Add(dk);
                        sm3.Add(dn);
                        sm3_e.Add(dn.Date + TimeSpan.Parse("7:00:00"));
                        sm1.Add(dn.Date + TimeSpan.Parse("7:00:00"));
                        sm1_e.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                    }

                }
                else if (dk > dn.Date + TimeSpan.Parse("23:59:59") && dk < dn.Date + TimeSpan.Parse("1.07:00:00"))
                {
                    string s = Convert.ToString(dn.Date + TimeSpan.Parse("1.00:00:00"));

                    if (!sm3.Contains(dn) && !sm3_e.Contains(dk))
                    {
                        sm3.Add(dn);
                        sm3_e.Add(dn.Date + TimeSpan.Parse("7:00:00"));
                        sm3.Add(dn.Date + TimeSpan.Parse("1.00:00:00"));
                        sm3_e.Add(dk);
                        sm2.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                        sm2_e.Add(dn.Date + TimeSpan.Parse("23:59:59"));
                        sm1.Add(dn.Date + TimeSpan.Parse("7:00:00"));
                        sm1_e.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                    }


                }
            }

            else if (dn < dn.Date + TimeSpan.Parse("15:30:00") && dn > dn.Date + TimeSpan.Parse("7:00:00")) // 1 смена
            {
                if (dk > dn.Date + TimeSpan.Parse("15:30:00") && dk < dn.Date + TimeSpan.Parse("23:59:59"))
                {
                    if (!sm3.Contains(dn) && !sm3_e.Contains(dk))
                    {
                        sm1.Add(dn);
                        sm1_e.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                        sm2.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                        sm2_e.Add(dk);

                    }

                }

                if (dk > dn.Date + TimeSpan.Parse("1.00:00:00") && dk < dn.Date + TimeSpan.Parse("1.07:00:00"))
                {
                    if (!sm3.Contains(dn) && !sm3_e.Contains(dk))
                    {
                        sm1.Add(dn);
                        sm1_e.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                        sm2.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                        sm2_e.Add(dn.Date + TimeSpan.Parse("23:59:59"));
                        sm3.Add(dn.Date + TimeSpan.Parse("1.00:00:00"));
                        sm3_e.Add(dk);

                    }

                }

            }
            else if (dn < dn.Date + TimeSpan.Parse("23:59:59") && dn > dn.Date + TimeSpan.Parse("15:30:00")) // 1 смена
            {
                if (dk > dn.Date + TimeSpan.Parse("1.00:00:00") && dk < dn.Date + TimeSpan.Parse("1.07:00:00"))
                {
                    if (!sm3.Contains(dn) && !sm3_e.Contains(dk))
                    {
                        sm2.Add(dn);
                        sm2_e.Add(dn.Date + TimeSpan.Parse("15:30:00"));
                        sm3.Add(dn.Date + TimeSpan.Parse("1.00:00:00"));
                        sm3_e.Add(dk);
                    }

                }
            }
        }


        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public void button4_Click(object sender, EventArgs e)
        {
            

            progressBar1.Value = 0;
            label30.Text = progressBar1.Value.ToString();
            DateTime beg_day, end_day;
            List<DateTime> begday_ar = new List<DateTime>();
            List<DateTime> endday_ar = new List<DateTime>();
            List<TimeSpan> ar_smen = new List<TimeSpan>();
            List<double> autoriz_sec = new List<double>();
            string[,] ar_day_time;
            List<TimeSpan> ar_hope = new List<TimeSpan>();
            TimeSpan oper_smen = TimeSpan.Parse("00:00:00");
            end_day = Convert.ToDateTime("11.11.2077");
                Excel.Application app = null;
                Excel.Workbook workbook = null;
                Excel.Worksheet worksheet = null;
                Excel.Worksheet worksheet2 = null;
            //Excel.Range workSheet_range = null;
            app = new Excel.Application();
                app.SheetsInNewWorkbook = 1;//обязательно до создания новой книги
                workbook = app.Workbooks.Add(1);
                workbook.Worksheets.Add();
                worksheet = workbook.Sheets[1];
                worksheet2 = workbook.Sheets[2];
                worksheet.Name = "Отчёт";
                worksheet2.Name = "Исходные";


            worksheet2.Cells[1, 1] = "Программа";
            worksheet2.Cells[1, 2] = "Обозначение";
            worksheet2.Cells[1, 3] = "Оператор";
            worksheet2.Cells[1, 4] = "Оборудование";
            worksheet2.Cells[1, 5] = "Начало";
            worksheet2.Cells[1, 6] = "Завершение";
            worksheet2.Cells[1, 7] = "Машинное время";
            worksheet2.Cells[1, 8] = "Норма времени";
            worksheet2.Cells[1, 9] = "Останов";
            worksheet2.Cells[1, 10] = "Пауза";
            worksheet2.Cells[1, 11] = "Счётчик";
            worksheet2.Cells[1, 12] = "Отклонение";
           // MessageBox.Show(dataar.Columns.Count.ToString() + " - " + dataar.Rows.Count.ToString());
            for (int j = 1; j < 13; j++)
            {
                for (int i = 2; i <= dt.Rows.Count+1; i++)
                {
                    //MessageBox.Show(dt.Rows[j - 1][i - 1].ToString());
                    worksheet2.Cells[i, j] = dt.Rows[i - 2][j - 1].ToString();
                    //worksheet.Cells[i, j] = dataar.Rows[i ][j ].ToString();
                }
            }

            worksheet2.get_Range("A1", "N" + dt.Rows.Count+3).Cells.Font.Size = 13;
            worksheet2.get_Range("A1", "N" + dt.Rows.Count+3).Columns.EntireColumn.AutoFit();



            if (radioButton9.Checked == true)
            {
                worksheet.Cells[1, 1].Value = "Оператор";
                worksheet.Cells[1, 2].Value = "Станок";
                worksheet.Cells[1, 3].Value = "День";
                worksheet.Cells[1, 4].Value = "Период";
                worksheet.Cells[1, 5].Value = "Переналадка";
                worksheet.Cells[1, 6].Value = "Отклонение";
                worksheet.Cells[1, 7].Value = "Длительность смены";
                worksheet.Cells[1, 8].Value = "Машинное время, ч.";
                worksheet.Cells[1, 9].Value = "Остановы, ч.";
                worksheet.Cells[1, 10].Value = "Пауза, ч.";
                worksheet.Cells[1, 11].Value = "КПД, %";
                worksheet.Cells[1, 12].Value = "Программы";
               // worksheet.Cells[1, 13].Value = "Инструмент";
               // worksheet.Cells[1, 14].Value = "Итог";
                
                var shapk = worksheet.get_Range("A1", "L1").Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                //worksheet.get_Range("A1", "J1").AutoFilter();

                // worksheet.get_Range("A1", "J1").Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("A1", "L1").Cells.Font.Size = 14;
                
                worksheet.get_Range("A1", "L1").Cells.Font.Bold = true;
                int k = 2;
                int rowcount = 0;

                


                foreach (string opert in comboBox2.Items)
                {
                    progressBar1.Value += (100/comboBox2.Items.Count);
                    label30.Text = progressBar1.Value.ToString();
                    int all = 0;
                    comboBox2.Text = opert;
                    repoday(opert);
                    beg_day = DateTime.MaxValue;
                    end_day = DateTime.MinValue;
                    double seconds = 0, sec = 0;
                    double col = 0;
                    ar_day_time = new string[2, dataGridView4.Rows.Count];

                    List<DateTime> l_day = new List<DateTime>();
                    List<DateTime> l_day_beg = new List<DateTime>();
                    List<DateTime> l_day_end = new List<DateTime>();
                    List<DateTime> cringe1 = new List<DateTime>();
                    List<DateTime> cringe2 = new List<DateTime>();
                    List<DateTime> unikend_ar = new List<DateTime>();
                    List<DateTime> unikbeg_ar = new List<DateTime>();
                    List<DateTime> smena1beg_ar = new List<DateTime>();
                    List<DateTime> smena1end_ar = new List<DateTime>();
                    List<DateTime> smena2beg_ar = new List<DateTime>();
                    List<DateTime> smena2end_ar = new List<DateTime>();
                    List<DateTime> smena3beg_ar = new List<DateTime>();
                    List<DateTime> smena3end_ar = new List<DateTime>();


                    List<int> l_kolvo = new List<int>();

                    // СОЗДАЁМ МАССИВ ДНЕЙ
                    try
                    {
                        l_day.Add(Convert.ToDateTime(dataGridView4["День", 0].Value.ToString()).Date);
                        for (int i = 0; i < dataGridView4.Rows.Count; i++)
                        {
                            bool day_is = false;
                            l_day_beg.Add(Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(0, 19)));
                            l_day_end.Add(Convert.ToDateTime(dataGridView4["Период", i].Value.ToString().Substring(22, 19)));
                            for (int y = 0; y < l_day.Count; y++)
                            {
                                if (l_day[y] == Convert.ToDateTime(dataGridView4["День", i].Value.ToString()).Date)
                                {
                                    day_is = true;
                                }
                            }
                            if (day_is == false)
                            {
                                l_day.Add(Convert.ToDateTime(dataGridView4["День", i].Value.ToString()));
                            }
                        }
                    }
                    catch { }
                    // МАССИВ ДНЕЙ ЗАПОЛНЕН

                    // РАБОТА
                    for (int day_ind = 0; day_ind < l_day.Count; day_ind++)
                    {
                        //int cnt = 0;
                        List<DateTime> iday_ar = new List<DateTime>();
                        List<DateTime> idayend_ar = new List<DateTime>();


                        // ЗАПОЛНЯЕМ МАССИВ НАЧАЛЬНЫХ ДАТ
                        for (int beg_ind = 0; beg_ind < l_day_beg.Count; beg_ind++)
                        {
                            if ((l_day[day_ind].Day == l_day_beg[beg_ind].Day) && (l_day[day_ind].Month == l_day_beg[beg_ind].Month) && (l_day[day_ind].Year == l_day_beg[beg_ind].Year))
                            {
                                //cnt++;
                                iday_ar.Add(l_day_beg[beg_ind]);
                                idayend_ar.Add(l_day_end[beg_ind]);
                            }
                            //
                        }
                        // МАССИВ ЗАПОЛНЕН

                        // ОПРЕДЕЛЕНИЕ ЗНАЧЕНИЙ НАЧАЛА СМЕН
                        DateTime beg_min = iday_ar[0];
                        DateTime fuckingday = iday_ar[0].Date;
                        DateTime beg_max = iday_ar[0];
                        for (int y = 0; y < iday_ar.Count; y++)
                        {
                            DateTime base_date = iday_ar[y].Date;
                            //(beg_min > iday_ar[y]) && 
                            if ((iday_ar[y] - base_date >= TimeSpan.Parse("0:00:00")) && (iday_ar[y] - base_date <= TimeSpan.Parse("14:30:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:00:00")))
                            {
                                smena1beg_ar.Add(iday_ar[y]);
                                smena1end_ar.Add(idayend_ar[y]);
                            }

                            if ((iday_ar[y] - base_date > TimeSpan.Parse("14:30:00")) && (iday_ar[y] - base_date <= TimeSpan.Parse("23:30:00")) && (idayend_ar[y] - base_date <= TimeSpan.Parse("3.23:50:00")))
                            {
                                smena2beg_ar.Add(iday_ar[y]);
                                smena2end_ar.Add(idayend_ar[y]);
                            }

                        }
                    }

                    for (int i = 0; i < smena1beg_ar.Count; i++)
                    {
                        for (int j = 0; j < smena2beg_ar.Count; j++)
                        {
                            try
                            {
                                if ((smena1beg_ar[i].Date == smena2beg_ar[j].Date) && (smena2beg_ar[j] - smena1beg_ar[i] < TimeSpan.Parse("3:00:00")) && (smena1end_ar[i] > smena2end_ar[j]))
                                {
                                    smena2beg_ar[j] = smena1beg_ar[i];
                                    smena1beg_ar.Remove(smena1beg_ar[i]);
                                    smena1end_ar.Remove(smena1end_ar[i]);
                                }
                            }
                            catch { }
                            
                        }
                    }

                    List<DateTime> list_begday = new List<DateTime>();
                    List<DateTime> list_endday = new List<DateTime>();
                    List<DateTime> list_begday2 = new List<DateTime>();
                    List<DateTime> list_endday2 = new List<DateTime>();
                    List<DateTime> list_begday3 = new List<DateTime>();
                    List<DateTime> list_endday3 = new List<DateTime>();

                    for (int day = 0; day < l_day.Count; day++)
                    {
                        DateTime beg_max = DateTime.MinValue;
                        DateTime beg_min = DateTime.MaxValue;
                        DateTime beg_max2 = DateTime.MinValue;
                        DateTime beg_min2 = DateTime.MaxValue;
                        DateTime beg_max3 = DateTime.MinValue;
                        DateTime beg_min3 = DateTime.MaxValue;
                        DateTime sbeg1 = DateTime.MinValue;
                        DateTime send1 = DateTime.MaxValue;

                        for (int i = 0; i < smena1beg_ar.Count; i++)
                        {
                            if (l_day[day].Date == smena1beg_ar[i].Date)
                            {
                                if (beg_min > smena1beg_ar[i])
                                {
                                    beg_min = smena1beg_ar[i];
                                }

                                if (beg_max < smena1end_ar[i])
                                {
                                    beg_max = smena1end_ar[i];
                                }

                                if (sbeg1 < smena1beg_ar[i])
                                {
                                    sbeg1 = smena1beg_ar[i];
                                }
                            }
                        }

                        for (int i = 0; i < smena2beg_ar.Count; i++)
                        {
                            if (l_day[day].Date == smena2beg_ar[i].Date)
                            {
                                if (beg_min2 > smena2beg_ar[i])
                                {
                                    beg_min2 = smena2beg_ar[i];
                                }

                                if (beg_max2 < smena2end_ar[i])
                                {
                                    beg_max2 = smena2end_ar[i];
                                }
                            }
                        }

                        for (int i = 0; i < smena3beg_ar.Count; i++)
                        {
                            if (l_day[day].Date == smena3beg_ar[i].Date)
                            {
                                if (beg_min3 > smena3beg_ar[i])
                                {
                                    beg_min3 = smena3beg_ar[i];
                                }

                                if (beg_max3 < smena3end_ar[i])
                                {
                                    beg_max3 = smena3end_ar[i];
                                }
                            }
                        }

                        if (beg_min.Date == l_day[day].Date)
                        {
                            list_begday.Add(beg_min);
                            list_endday.Add(beg_max);
                        }

                        if (beg_min2.Date == l_day[day].Date)
                        {
                            list_begday2.Add(beg_min2);
                            list_endday2.Add(beg_max2);
                        }

                        if (beg_min3.Date == l_day[day].Date)
                        {
                            list_begday3.Add(beg_min3);
                            list_endday3.Add(beg_max3);
                        }
                    }


                    TimeSpan smenatime = TimeSpan.Parse("00:00:00");

                    for (int i = 0; i < list_begday.Count; i++)
                    {
                        smenatime += list_endday[i] - list_begday[i];
                    }

                    for (int i = 0; i < list_begday2.Count; i++)
                    {
                        smenatime += list_endday2[i] - list_begday2[i];
                    }

                    for (int i = 0; i < list_begday3.Count; i++)
                    {
                        smenatime += list_endday3[i] - list_begday3[i];
                    }

                    ar_smen.Add(smenatime);


                    List<string> prostoi_done = new List<string>(); 
                    
                    List<string> all_prostoi_done = new List<string>();


                    for (int u = 0; u < dataGridView4.Rows.Count; u++)
                    {
                        string prost = "";
                        TimeSpan all_prost = TimeSpan.Zero ;
                        
                        for (int t = 0; t < listBox7.Items.Count; t++)
                        {
                            if ((dataGridView4[1,u].Value.ToString() == listBox7.Items[t].ToString().Split('|')[0]) && (Convert.ToDateTime(dataGridView4[3, u].Value.ToString().Split('-')[0].Trim()) <= Convert.ToDateTime(listBox7.Items[t].ToString().Split('|')[1])) && (Convert.ToDateTime(dataGridView4[3, u].Value.ToString().Split('-')[1].Trim()) >= Convert.ToDateTime(listBox7.Items[t].ToString().Split('|')[2])))
                            {
                                prost += "["+listBox7.Items[t].ToString().Split('|')[3].Replace(" - ", " = ") + "], ";
                            }
                        }

                        prostoi_done.Add(prost);

                    }

                   

                    for (int i = 0; i < dataGridView4.Rows.Count; i++)
                    {
                        rowcount++;
                            all++;
                            worksheet.Cells[i + k, 1].Value = opert;
                            worksheet.Cells[i + k, 2].Value = dataGridView4[1, i].Value.ToString();
                            worksheet.Cells[i + k, 3].Value = Convert.ToDateTime(dataGridView4[2, i].Value.ToString());
                            worksheet.Cells[i + k, 4].Value = dataGridView4[3, i].Value.ToString();
                            worksheet.Cells[i + k, 5].Value = prostoi_done[i];
                            worksheet.Cells[i + k, 6].Value = dataGridView4[6, i].Value.ToString();
                            worksheet.Cells[i + k, 7].Value = dataGridView4[4, i].Value.ToString();
                            worksheet.Cells[i + k, 8].Value = Convert.ToDouble(dataGridView4[5, i].Value.ToString());
                            worksheet.Cells[i + k, 9].Value = Convert.ToDouble(dataGridView4[7, i].Value.ToString());
                            worksheet.Cells[i + k, 10].Value = Convert.ToDouble(dataGridView4[8, i].Value.ToString());
                            worksheet.Cells[i + k, 11].Value = Convert.ToDouble(dataGridView4[9, i].Value.ToString());
                            worksheet.Cells[i + k, 11].Font.Color = Color.Red;
                            worksheet.Cells[i + k, 12].Value = dataGridView4[10, i].Value.ToString();
                            
                            if (worksheet.Cells[i + k, 11].Value >= Convert.ToDouble(50))
                            {
                                worksheet.Cells[i + k, 11].Font.Color = Color.Green;
                            }
                            if (beg_day > Convert.ToDateTime(dataGridView4[2, i].Value))
                            {
                                beg_day = Convert.ToDateTime(dataGridView4[2, i].Value);
                            }
                            if (end_day < Convert.ToDateTime(dataGridView4[2, i].Value))
                            {
                                end_day = Convert.ToDateTime(dataGridView4[2, i].Value);
                            }
                        

                    }
                    k += dataGridView4.Rows.Count + 1;
                    begday_ar.Add(Convert.ToDateTime(beg_day.ToString()));
                    endday_ar.Add(Convert.ToDateTime(end_day.ToString()));



                }

                List<int> bid = new List<int>();
                List<int> eid = new List<int>();
                List<int> emptyxls_id = new List<int>();
                bid.Add(2);
                for (int i = 2; i < k; i++)
                {
                    if ((worksheet.Cells[i, 11].Value is null))
                    {
                        if ((worksheet.Cells[i + 1, 11].Value != null))
                        {
                            bid.Add(i + 1);
                        }

                        eid.Add(i - 1);
                    }

                }

                int rowcount2 = 0;

                for (int i = 0; i < bid.Count; i++)
                {

                    for (int u = 2; u < k; u++)
                    {
                        if (worksheet.Cells[u, 11].Value == null)
                        {

                            double sum_MASH = app.WorksheetFunction.Sum(worksheet.get_Range("H" + bid[i], "H" + eid[i]));
                            double sum_OST = app.WorksheetFunction.Sum(worksheet.get_Range("I" + bid[i], "I" + eid[i]));
                            double sum_PAUSE = app.WorksheetFunction.Sum(worksheet.get_Range("J" + bid[i], "J" + eid[i]));

                            rowcount2++;

                            worksheet.Cells[u, 11].Value = Math.Round((sum_MASH / (sum_MASH + sum_OST + sum_PAUSE)) * 100, 2);
                            worksheet.Cells[u, 11].Interior.Color = Color.Red;
                            worksheet.Cells[u, 11].Font.Color = Color.White;
                            worksheet.Cells[u, 11].Font.Bold = true;
                            worksheet.Cells[u, 10].Font.Bold = true;
                            worksheet.Cells[u, 9].Font.Bold = true;
                            worksheet.Cells[u, 8].Font.Bold = true;
                            worksheet.Cells[u, 7].Font.Bold = true;
                            worksheet.Cells[u, 6].Font.Bold = true;
                            worksheet.Cells[u, 5].Font.Bold = true;
                            worksheet.Cells[u, 4].Font.Bold = true;
                            worksheet.Cells[u, 7].Font.Bold = true;
                            worksheet.Cells[u, 1].Font.Bold = true;
                            try
                            {
                                if (worksheet.Cells[u, 11].Value >= Convert.ToDouble(50)) { worksheet.Cells[u, 11].Interior.Color = Color.Green; }
                                double prog = app.WorksheetFunction.Average(worksheet.get_Range("H" + bid[i], "H" + eid[i]));
                                worksheet.Cells[u, 8].Value = Math.Round(prog, 2);// + " - " + Math.Round(sum_MASH, 2);
                                double ost = app.WorksheetFunction.Average(worksheet.get_Range("I" + bid[i], "I" + eid[i]));
                                worksheet.Cells[u, 9].Value = Math.Round(ost, 2);// + " - " + Math.Round(sum_OST, 2);
                                double pause = app.WorksheetFunction.Average(worksheet.get_Range("J" + bid[i], "J" + eid[i]));
                                worksheet.Cells[u, 10].Value = Math.Round(pause, 2);// + " - " + Math.Round(sum_PAUSE, 2);
                            }
                            catch (COMException exception)
                            {
                               //
                            }
                            
                            worksheet.Cells[u, 4].Value = begday_ar[i].ToString().Substring(0, 10) + " - " + endday_ar[i].ToString().Substring(0, 10);
                            worksheet.Cells[u, 1].Value = worksheet.Cells[u-1, 1].Value;
                            emptyxls_id.Add(u);

                            break;
                        }
                    }

                }

                for (int i = 0; i < emptyxls_id.Count; i++)
                {
                    if (ar_smen[i].ToString().Length > 8)
                    {
                        string itog = ar_smen[i].ToString();
                        string days = itog.Split('.')[0];
                        string hours = itog.Split('.')[1].Split(':')[0];
                        int int_hours = Convert.ToInt32(Convert.ToInt32(days) * 24) + Convert.ToInt32(hours);
                        string str_itog = int_hours.ToString() + ':' + itog.Split('.')[1].Split(':')[1] + ':' + itog.Split('.')[1].Split(':')[2];
                        worksheet.Cells[emptyxls_id[i], 7].Value = str_itog;
                    }
                    else
                    {
                        worksheet.Cells[emptyxls_id[i], 7].Value = ar_smen[i].ToString();
                    }

                }

                int rowsexcel = rowcount2 + rowcount;
                int otstup = 6;

                List<string> legend_itog = new List<string>();

                for (int r = 0; r < legend.Count; r++) 
                { 
                    if ((legend[r].Split('-')[0].Trim() != "O0") && (legend[r].Split('-')[1].Trim() != "")) 
                    {
                        //legend.Remove(legend[r]); MessageBox.Show(legend[r]); 
                        legend_itog.Add(legend[r]);
                    }  
                }

                for (int r = 0; r < legend_itog.Count; r++) { worksheet.Cells[otstup + rowsexcel + r, 1].Value = legend_itog[r].ToString(); }
                    


                var cells = worksheet.get_Range("A1", "L" + (k - 1));
                cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние вертикальные
                cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // внутренние горизонтальные            
                cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // верхняя внешняя
                cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // правая внешняя
                cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous; // левая внешняя
                cells.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                worksheet.get_Range("A2", "L" + (k - 1)).Cells.Font.Size = 13;
                worksheet.get_Range("K1", "K" + (k - 1)).Cells.Font.Bold = true;
                cells.Columns.EntireColumn.AutoFit();
                cells = null;

            }
        
            worksheet.get_Range("A1", "L1").Cells.Style.WrapText = true;

            string fullfilename2003;
            fullfilename2003 = Microsoft.VisualBasic.Interaction.InputBox("Введите имя отчёта", "Имя отчёта", "Отчёт");
            if (System.IO.File.Exists(fullfilename2003)) System.IO.File.Delete(fullfilename2003);
            string month = fullfilename2003.Substring(fullfilename2003.Length - 3).Trim();
            switch (month)
            {
                case "янв":
                    month = "Январь";
                    break;
                case "фев":
                    month = "Февраль";
                    break;
                case "мар":
                    month = "Март";
                    break;
                case "апр":
                    month = "Апрель";
                    break;
                case "май":
                    month = "Май";
                    break;
                case "июн":
                    month = "Июнь";
                    break;
                case "июл":
                    month = "Июль";
                    break;
                case "aвг":
                    month = "Август";
                    break;
                case "сен":
                    month = "Сентябрь";
                    break;
                case "окт":
                    month = "Октябрь";
                    break;
                case "ноя":
                    month = "Ноябрь";
                    break;
                case "дек":
                    month = "Декабрь";
                    break;
            }
            
            string path = "\\\\Files\\обмен\\Отдел главного технолога\\Отчёты КПД оборудования\\" + month + @"\";
                workbook.SaveAs(path + fullfilename2003 + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault); //формат Excel 2007+

                workbook.Close(0); //false - закрыть рабочую книгу не сохраняя изменения

            int tExcelPID = 0;
            int tHwnd = 0;
            tHwnd = app.Hwnd; //Получим HWND окна
            System.Diagnostics.Process tExcelProcess;
            progressBar1.Value = 100;
            label30.Text = progressBar1.Value.ToString();
            GetWindowThreadProcessId((IntPtr)tHwnd, out tExcelPID); //По HWND получим PID
            tExcelProcess = System.Diagnostics.Process.GetProcessById(tExcelPID); //Подключимся к процессу
                                                                                  ////Убийство процесса Excel
            tExcelProcess.Kill();
            tExcelProcess = null;
               
            progressBar1.Value = 0;
            label30.Text = progressBar1.Value.ToString();
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
            dt.Columns.Add("Норма времени", typeof(String));
            dt.Columns.Add("Остановы, ч.", typeof(String));
            dt.Columns.Add("Пауза, ч.", typeof(String));
            dt.Columns.Add("Счётчик", typeof(String));
            dt.Columns.Add("Отклонение", typeof(TimeSpan));
            dt.Columns.Add("Числитель, с.", typeof(int));
            dt.Columns.Add("Знаменатель, с.", typeof(int));
            dataGridView6.Columns.Add("День", "День");
            dataGridView6.Columns.Add("1 станок", "1 станок");
            dataGridView6.Columns.Add("2 станок", "2 станок");
            dataGridView6.Columns.Add("3 станок", "3 станок");
            dataGridView6.Columns.Add("4 станок", "4 станок");
            dataGridView6.Columns.Add("5 станок", "5 станок");
            dataGridView6.Columns.Add("6 станок", "6 станок");
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

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }
    }
}
