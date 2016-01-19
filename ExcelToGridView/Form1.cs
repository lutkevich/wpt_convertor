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

namespace ExcelToGridView
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();  
        }

        private void button1_Click(object sender, EventArgs e)
        {
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
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    columnNames[0] = dt.Columns[i].ColumnName;
                }                             
                
                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                
                dataGridView1.DataSource = dt;
                app.Quit();
            }
            else
                Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                String constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                                ofd.FileName +
                                ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

                System.Data.OleDb.OleDbConnection con =
                    new System.Data.OleDb.OleDbConnection(constr);
                con.Open();

                DataSet ds = new DataSet();
                DataTable schemaTable = con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables,
                    new object[] { null, null, null, "TABLE" });

                string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                string select = String.Format("SELECT * FROM [{0}]", sheet1);

                System.Data.OleDb.OleDbDataAdapter ad = 
                    new System.Data.OleDb.OleDbDataAdapter(select, con);

                ad.Fill(ds);

                DataTable tb = ds.Tables[0];
                con.Close();
                dataGridView1.DataSource = tb;
                con.Close();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл для открытия",
                        "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //http://exceldatareader.codeplex.com/

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;

                System.IO.FileStream stream = 
                    System.IO.File.Open(ofd.FileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);

                Excel.IExcelDataReader IEDR;

                int fileformat = ofd.SafeFileName.IndexOf(".xlsx");

                if (fileformat > -1)
                {
                    //2007 format *.xlsx
                    IEDR = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);

                }
                else
                {
                    //97-2003 format *.xls
                    IEDR = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
                }

                //Если данное значение установлено в true
                //то первая строка используется в качестве 
                //заголовков для колонок
                IEDR.IsFirstRowAsColumnNames = true;

                DataSet ds = IEDR.AsDataSet();

                //Устанавливаем в качестве источника данных dataset 
                //с указанием номера таблицы. Номер таблицы указавает 
                //на соответствующий лист в файле нумерация листов 
                //начинается с нуля.
                dataGridView1.DataSource = ds.Tables[0];
                IEDR.Close();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл для открытия",
                 "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Переменная указывающая, использовать ли
            //первую строку, в качестве заголовков для 
            //столбцов
            bool IsFirstRowAsColumnNames = true;

            char[] rowSplitter = { '\r', '\n' };

            IDataObject dataInClipboard = Clipboard.GetDataObject();

            string stringInClipboard = 
                (string)dataInClipboard.GetData(DataFormats.Text);

            string[] rowsInClipboard = 
                stringInClipboard.Split(rowSplitter, 
                StringSplitOptions.RemoveEmptyEntries);
            DataTable dt = new DataTable();

            if (IsFirstRowAsColumnNames)
            {
                //В данном случае берется первая строка 
                //и разделятся на отдельные части с помощью 
                //класса Split. Данные помещаются в массив
                //после чего, циклом forech добавляются 
                //заголовки столбцов
                string[] words = rowsInClipboard[0].Split('\t');
                foreach (string word in words)
                {
                    dt.Columns.Add(word);
                }

                for (int i = 2; i <= rowsInClipboard.Length; i++)
                {
                    string[] rows = rowsInClipboard[i - 1].Split('\t');
                    dt.Rows.Add(rows);
                }
            }
            else
            {
                //Если в скопированных данных отсутствуют 
                //заголовки для столбцов, то 
                // автоматически пронумеровать их
                for (int colc = 1; colc <= rowsInClipboard[0].Split('\t').Length; colc++)
                {
                    dt.Columns.Add("Столбец " + colc);
                }

                for (int i = 1; i <= rowsInClipboard.Length; i++)
                {
                    string[] rows = rowsInClipboard[i - 1].Split('\t');
                    dt.Rows.Add(rows);
                }
            }
            dataGridView1.DataSource = dt;   
        }
    }
}
