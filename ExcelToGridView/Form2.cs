using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using ExcelObj = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;


namespace ExcelToGridView
{

    public partial class Form1 : Form
    {

        string[] linewpt = new string[1000]; //массив строк в wpt
        double[] wpt_lat = new double[1000];//широта
        double[] wpt_long= new double[1000];//долгота
        string[] wpt_name= new string[1000];//имя
        int[] wpt_mark = new int[1000];//баллы
        string[] wpt_desc= new string[1000];//описание

        double numlat, numlong;
        int maxrow; //максимальное число столбцов
        int maxcol; //максимальное число строк
        int max_mark; //минимальный балл
        int min_mark; //максимальный балл
        int mid_mark; //диапазон цвета
        public Form1()
        {
            InitializeComponent();  
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            string[,] mydata = new string[100, 100]; //массив загруженный их excel
            string[,] mydataout = new string[100, 100];//массив выгружаемых wpt 
            double numlat, numlong;
            maxrow = 0; 
            maxcol = 0; 
             
            ExcelObj.Application app = new ExcelObj.Application();
            
            OpenFileDialog ofd = new OpenFileDialog();
            //Задаем расширение имени файла по умолчанию.
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, которая определяет
            //варианты, доступные в поле "Файлы типа" диалогового
            //окна.
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            //Задаем заголовок диалогового окна.
            ofd.Title = "Выберите документ для загрузки данных";

            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            ExcelObj.Workbook workbook;
            DataTable dt = new DataTable();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                
                //Устанавливаем номер листа из котрого будут извлекаться данные
                //Листы нумеруются от 1
                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;
                for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(new DataColumn((ShtRange.Cells[1, Cnum] as ExcelObj.Range).Value2.ToString()));
                }
                dt.AcceptChanges();

                string[] columnNames = new String[dt.Columns.Count];
                maxcol = dt.Columns.Count;
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }
                maxrow = ShtRange.Rows.Count;
                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                            mydata [Rnum, Cnum] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                           // data_.SetValue(dr.Rows(iRow).Item(iCol).ToString, {iRow, iCol})
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                for (int i = 2; i <=  maxrow; i++)
                {
                    //======== конвертирует lat =========
                    try
                    {
                        wpt_lat[i] = Convert.ToDouble(mydata[i, 2]) + Convert.ToDouble(mydata[i, 3]) / 60; 
                    }
                    catch (FormatException e1)
                    {
                        MessageBox.Show("Широта имеет некорректное значение \n" + "Cтрока: " + i.ToString() + "\n Значение: " + mydata[i, 2] + mydata[i, 3]);
                        wpt_lat[i] = 0;
                    }
                    //======== конвертирует long =========
                    try
                    { 
                        wpt_long[i] = Convert.ToDouble(mydata[i, 4]) + Convert.ToDouble(mydata[i, 5]) / 60; 
                    }
                    catch (FormatException e1)
                    {
                        MessageBox.Show("Долгота имеет некорректное значение \n" + "Cтрока: " + i.ToString() + "\n Значение: " + mydata[i, 4] + mydata[i, 5]);
                        wpt_long[i] = 0;
                    }
                    wpt_lat[i] = Math.Round(wpt_lat[i], 6);
                    wpt_long[i] = Math.Round(wpt_long[i], 6);
                    wpt_name[i]=mydata[i, 1];
                    wpt_desc[i]=mydata[i, 7];
                    try
                    {
                        wpt_mark[i] = Convert.ToInt32(mydata[i, 6]);
                        
                    }
                    catch (FormatException e1)
                    {
                        MessageBox.Show("Бальность точки не целое число \n"+"Cтрока: " + i.ToString() + "\n Значение: " + mydata[i, 6]);
                    }
                    if (i == 2) { max_mark = wpt_mark[i]; min_mark = wpt_mark[i]; }
                    if (max_mark < wpt_mark[i]) { max_mark = wpt_mark[i]; }
                    if (min_mark > wpt_mark[i]) { min_mark = wpt_mark[i]; }
                    //linewpt[i] = (i - 1).ToString() + ", " + mydata[i, 1] + ", " + numlat.ToString().Replace(",", ".") + ", " + numlong.ToString().Replace(",", ".") + "0 ,0,0,0,3,0,65535," + mydata[i, 7] + ",0,0,0,-777, 6, 0,17,0,10.0,2,,,\n";
                }

                dataGridView1.DataSource = dt;
                //================ выводим статистику ====================
                textBox3.Text = "Всего точек: " + (maxrow - 1).ToString() + "\n Максимальный балл:" + max_mark.ToString() + "\n Минимальный балл:" + min_mark.ToString();
                mid_mark = (max_mark - min_mark) / 4; //4 диапазона цветов
                trackBar1.Minimum = min_mark; trackBar1.Maximum = max_mark; trackBar1.Value = min_mark + mid_mark*1;
                trackBar2.Minimum = min_mark; trackBar2.Maximum = max_mark; trackBar2.Value = min_mark + mid_mark * 2;
                trackBar3.Minimum = min_mark; trackBar3.Maximum = max_mark; trackBar3.Value = min_mark + mid_mark * 3;
                trackBar4.Minimum = min_mark; trackBar4.Maximum = max_mark; trackBar4.Value = min_mark + mid_mark * 4;
                app.Quit();
            }
            else
                Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string textout;
            string wpt_name_out;
            string wpt_lat_out;
            string wpt_long_out;
            string wpt_mark_out;
            string wpt_desc_out;
            string wpt_color_out;

            const string wpt_red = "255", wpt_yellow = "65535", wpt_blue = "16776960", wpt_green = "65280";

            textout = "OziExplorer Waypoint File Version 1.1\n";
            textout = textout + "WGS 84\n";
            textout= textout + "Reserved 2\n";
            textout=textout +"garmin\n";
            for (int i = 2; i <= maxrow; i++)
            {

                wpt_name_out = wpt_name[i];
                if (checkBox1.Checked = true) //добавить балл в назавние
                {
                    wpt_name_out = wpt_name_out + "-" + wpt_mark[i];
                }
                wpt_lat_out = wpt_lat[i].ToString().Replace(",", ".");
                wpt_long_out = wpt_long[i].ToString().Replace(",", ".");
                wpt_desc_out = wpt_desc[i];
                wpt_color_out = "65535";
                if (checkBox2.Checked = true) //раскрасить точки
                {
                    if (wpt_mark[i] < min_mark + mid_mark * 1) { wpt_color_out = wpt_green; }
                    if ((wpt_mark[i] >= min_mark + mid_mark * 1) &&  (wpt_mark[i] < min_mark + mid_mark * 2) ){ wpt_color_out = wpt_blue; }
                    if ((wpt_mark[i] >= min_mark + mid_mark * 2) && (wpt_mark[i] < min_mark + mid_mark * 3)) { wpt_color_out = wpt_yellow; }
                    if (wpt_mark[i] >= min_mark + mid_mark * 4)  { wpt_color_out = wpt_red; }
                }

                

                linewpt[i] = (i - 1).ToString() + ", " + wpt_name_out + ", " + wpt_lat_out + ", " + wpt_long_out + "0 ,0,0,0,3,0," + wpt_color_out + "," + wpt_desc_out + ",0,0,0,-777, 6, 0,17,0,10.0,2,,,\n";
                             textout = textout + linewpt[i];
            }


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "OZI waypoint (*.wpt)|*.wpt|Все файлы (*.*)|*.*";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                StreamWriter streamWriter = new StreamWriter(saveFileDialog.FileName);
                streamWriter.WriteLine(textout);
                streamWriter.Close();
            }


        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you want to exit?", "Vnedorog waypoint convertor",  MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                // Cancel the Closing event from closing the form.
                Application.Exit();         
            }
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*Form f2 = new Form();
            f2.Show();*/
            //AboutForm AbForm = new AboutForm();
            AboutBox1 abtbtn = new AboutBox1();
            abtbtn.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.vnedorog.ru");
        }
    }
}
