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
using CSharpJExcel;
using CSharpJExcel.Jxl;


namespace ExcelToGridView
{

    public partial class Form1 : Form
    {
        private string[] linewpt; //массив строк в wpt
        private double[] wpt_lat;//широта
        private double[] wpt_long;//долгота
        private string[] wpt_name;//имя
        private int[] wpt_mark;//баллы
        private string[] wpt_desc;//описание

        double numlat, numlong;
        int max_mark; //минимальный балл
        int min_mark; //максимальный балл
        int mid_mark; //диапазон цвета

        int numrows;//максимальное число строк
        int numcols;//максимальное число столбцов

         
        private bool[] wpt_error;
        private string[,] excel_data;
        private string[,] mydata;
        private string[,] mydataout;
        public Form1()
        {
            InitializeComponent();  
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //---------- создаем таблицу для вывода в GridView ---------------
            DataTable table; //таблица для вывода в GridView
            string error_message;
            bool error_flag;

            error_message = "";
            error_flag = false;
            label4.Text = "0";
            textBox3.Text="";

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";//Задаем расширение имени файла по умолчанию.
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";//Задаем заголовок диалогового окна.

            //---------------открываем файл ----------
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
                try //проверяем можно ли открыть Excel файл
                {
                    Workbook.getWorkbook(new System.IO.FileInfo(ofd.FileName));
                }
                catch
                {
                    MessageBox.Show("Невозможно открыть файл! Возможно он открыйт в Excele...закройте его там");
                    return;
                }
                Workbook workbook1 = Workbook.getWorkbook(new System.IO.FileInfo(ofd.FileName));

                var sheet1 = workbook1.getSheet(0);
                var headerCells = sheet1.getRow(0);

                numrows = sheet1.getRows();
                numcols = sheet1.getColumns();
                //---------- инициализируем основыне массивы -------------
                linewpt = new string[numrows + 1]; //массив строк в wpt
                wpt_lat = new double[numrows + 1];//широта
                wpt_long = new double[numrows + 1];//долгота
                wpt_name = new string[numrows + 1];//имя
                wpt_mark = new int[numrows + 1];//баллы
                wpt_desc = new string[numrows + 1];//описание*/
                wpt_error = new bool[numrows + 1];// массив статусов успешной конвертации точек

                excel_data = new string[numcols+1, numrows+1]; //массив данных полученный из экселя: колонка, строка
                mydata = new string[numrows+1, numcols+1]; //массив загруженный их excel, порядок другой!!! строка - колонка 
                mydataout = new string[numrows + 1, numcols + 1];//массив выгружаемых wpt, порядок другой!!! строка - колонка  

                 //----------- читаем из Excel -------
                for (int i = 0; i < numrows ; i++)
                {
                    for (int j = 0; j < numcols; j++)
                    {
                        excel_data[j,i] = sheet1.getCell(j, i).getContents();
                        mydata[i, j] = excel_data[j, i];
                    }
                }
                workbook1.close();
               //------ формируем таблицу для GridView
                table = new DataTable();
                for (int j = 0; j < numcols; j++)
                {
                    table.Columns.Add(excel_data[j, 0], typeof(string)); 
                }
                //table.Columns.Add("Результат", typeof(string));
                for (int i = 1; i < numrows; i++)
                {
                    table.Rows.Add(excel_data[0, i], excel_data[1, i], excel_data[2, i], excel_data[3, i], excel_data[4, i], excel_data[5, i], excel_data[6, i], excel_data[7, i], excel_data[8, i]);
                }
                table.AcceptChanges();
                dataGridView1.DataSource = table;
                //================================================================================
                //----------------- конвертируем строки в значения -------------------------------
                for (int i = 1; i < numrows; i++)
                {
                    //======== конвертирует lat =========
                    if (mydata[i, 0] == "")
                    {
                        error_message = error_message + Environment.NewLine + "Нет названия точки, пропуск всей строки" + i.ToString();
                        error_flag = true;
                        wpt_name[i] = "";//имя точки
                        wpt_error[i] = true;
                    }
                    try
                    {
                        wpt_lat[i] = Convert.ToDouble(mydata[i, 1]) + Convert.ToDouble(mydata[i, 2]) / 60; 
                    }
                    catch (FormatException e1)
                    {
                        //MessageBox.Show("Широта, некорректное значение \n" + "Cтрока: " + i.ToString() + "\n Значение: " + mydata[i, 1] + mydata[i, 2]);
                        error_message = error_message + Environment.NewLine + "Широта имеет некорректное значение " + "Cтрока: " + i.ToString() + " Значение: " + mydata[i, 3];
                        wpt_lat[i] = 0;
                        error_flag = true;
                        wpt_error[i] = true;
                    }
                    //======== конвертирует long =========
                    try
                    { 
                        wpt_long[i] = Convert.ToDouble(mydata[i, 3]) + Convert.ToDouble(mydata[i, 4]) / 60; 
                    }
                    catch (FormatException e1)
                    {
                        //MessageBox.Show("Долгота, некорректное значение \n" + "Cтрока: " + i.ToString() + "\n Значение: " + mydata[i, 3] + mydata[i, 4]);
                        error_message = error_message + Environment.NewLine + "Долгота имеет некорректное значение " + "Cтрока: " + i.ToString() + " Значение: " + mydata[i, 3];
                        wpt_long[i] = 0;
                        error_flag = true;
                        wpt_error[i] = true;
                    }
                    wpt_lat[i] = Math.Round(wpt_lat[i], 6);//широта
                    wpt_long[i] = Math.Round(wpt_long[i], 6);//долгота
                    wpt_name[i]=mydata[i, 0];//имя точки
                    wpt_desc[i]=mydata[i, 6];//примечание
                    if (mydata[i, 5] != "" ) //если полее балл не пустое 
                    { 
                        try
                        {
                            wpt_mark[i] = Convert.ToInt32(mydata[i, 5]);//баллы
                        
                        }
                        catch (FormatException e1)
                        {
                            //MessageBox.Show("Бальность точки не целое число \n"+"Cтрока: " + i.ToString() + "\n Значение: " + mydata[i, 5]);
                            error_message = error_message + Environment.NewLine + "Бальность точки не целое число " + "Cтрока: " + i.ToString() + " Значение: " + mydata[i, 5];
                            error_flag = true;
                        }
                    }
                    if (i == 1) { max_mark = wpt_mark[i]; min_mark = wpt_mark[i]; }
                    if (max_mark < wpt_mark[i]) { max_mark = wpt_mark[i]; }
                    if (min_mark > wpt_mark[i]) { min_mark = wpt_mark[i]; }
                }
                //dataGridView1.DataSource = dt;
                //================ выводим статистику ====================
                
                label16.Text = (numrows - 1).ToString();//Всего точек:
                label17.Text = min_mark.ToString(); //Минимальный балл:
                label18.Text = max_mark.ToString(); //Максимальный балл:


                textBox3.ScrollBars = ScrollBars.Vertical;
                //textBox3.ForeColor = "red";
                textBox3.Text = error_message;

                
                if (error_flag == true)
                {
                    MessageBox.Show("В процессе конвертирования были обнаружены ошибки в исходном файле. См поле Ошибки");
                    label4.Text = "ОШИБКИ!!!";
                }

                mid_mark = (max_mark - min_mark) / 4; //4 диапазона цветов
                label8.Text = (min_mark + mid_mark * 0).ToString() + " - " + (min_mark + mid_mark * 1).ToString();
                label9.Text = (min_mark + mid_mark * 1 + 1).ToString() + " - " + (min_mark + mid_mark * 2).ToString();
                label10.Text = (min_mark + mid_mark * 2 + 1).ToString() + " - " + (min_mark + mid_mark * 3).ToString();
                label11.Text = (min_mark + mid_mark * 3 + 1).ToString() + " - " + (min_mark + mid_mark * 4).ToString();
                

               // app.Quit();
               
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

            textout = "OziExplorer Waypoint File Version 1.1" + Environment.NewLine;
            textout = textout + "WGS 84" + Environment.NewLine;
            textout = textout + "Reserved 2" + Environment.NewLine;
            textout = textout + "garmin" + Environment.NewLine;
            for (int i = 1; i < numrows; i++)
            {
                //if (wpt_name[i] != null)
                if ( wpt_error[i]  == false) // если на этой строке не было ошибки в конвертации
                {
                    wpt_name_out = wpt_name[i];
                    if (checkBox1.Checked == true) //добавить балл в назавние
                    {
                    wpt_name_out = wpt_name_out + "-" + wpt_mark[i];
                    }
                    wpt_lat_out = wpt_lat[i].ToString().Replace(",", ".");
                    wpt_long_out = wpt_long[i].ToString().Replace(",", ".");
                    wpt_desc_out = wpt_desc[i];
                    wpt_color_out = "65535";
                    if (checkBox2.Checked == true) //раскрасить точки
                    {
                        if (wpt_mark[i] < min_mark + mid_mark * 1) { wpt_color_out = wpt_green; }
                        if ((wpt_mark[i] >= min_mark + mid_mark * 1) &&  (wpt_mark[i] < min_mark + mid_mark * 2) ){ wpt_color_out = wpt_blue; }
                        if ((wpt_mark[i] >= min_mark + mid_mark * 2) && (wpt_mark[i] < min_mark + mid_mark * 3)) { wpt_color_out = wpt_yellow; }
                        if (wpt_mark[i] >= min_mark + mid_mark * 4)  { wpt_color_out = wpt_red; }
                    }

                    linewpt[i] = (i - 1).ToString() + ", " + wpt_name_out + ", " + wpt_lat_out + ", " + wpt_long_out + "0 ,0,0,0,3,0," + wpt_color_out + "," + wpt_desc_out + ",0,0,0,-777, 6, 0,17,0,10.0,2,,," + Environment.NewLine;
                    textout = textout + linewpt[i];
                }
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

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("В текущей версии доступен только импорт из Excel");
        }

        private void comboBox2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("В текущей версии доступен только экспорт в OZI Explorer (wpt)");
        }
    }
}
