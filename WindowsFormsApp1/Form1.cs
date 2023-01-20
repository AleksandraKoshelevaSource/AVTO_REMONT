using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


namespace WindowsFormsApp1
{

    public partial class F_Menu : System.Windows.Forms.Form
    {
        private Excel.Application excel_app;

        public F_Menu()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void Exit(object sender, EventArgs e)
        {
            Close();
        }

        private void OpenSotr(object sender, EventArgs e)
        {
            FormSotr f1 = new FormSotr();

            f1.ShowDialog();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void OpenPriceList(object sender, EventArgs e)
        {
            FormUsl f1 = new FormUsl();
            f1.ShowDialog();
        }

        private void ремонтToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormRemontAvto fr = new FormRemontAvto();
            fr.ShowDialog();
        }

        private void автомобилиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormRedAvto f = new FormRedAvto();
            f.ShowDialog();
        }

        private void списокСотрудниковToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //создать отчет в excell
            excel_app = new Excel.Application();
            excel_app.Visible = true;
            excel_app.SheetsInNewWorkbook = 1;
            excel_app.Workbooks.Add(Type.Missing);

            Excel.Range _excelCells = (Excel.Range)excel_app.get_Range("A1", "C1").Cells;
            _excelCells.Merge(Type.Missing);

            excel_app.Cells[1, 1].Value = "Список сотрудников на " + DateTime.Now;
            excel_app.Cells[1, 1].Font.Bold = true;
            excel_app.Cells[1, 1].Font.Size = 16;
            excel_app.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            excel_app.Cells[2, 1].Value = "№";
            excel_app.Columns[1].columnwidth = 6;

            excel_app.Cells[2, 2].Value = "ФИО сотрудника";
            excel_app.Columns[2].columnwidth = 30;

            excel_app.Cells[2, 3].Value = "Должность";
            excel_app.Columns[3].columnwidth = 30;

            for (int i = 1; i <= 3; i++)
            {
                excel_app.Cells[2, i].Font.Size = 14;
                excel_app.Cells[2, i].Font.Italic = true;
                excel_app.Cells[2, i].Font.Bold = true;
                excel_app.Cells[2, i].Borders.LineStyle = 1;
                excel_app.Cells[2, i].Borders.Weight = Excel.XlBorderWeight.xlThick;
                excel_app.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }

            string SQL_text = "SELECT * FROM MASTER";
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            SqlCommand comm = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = comm.ExecuteReader();
            int j = 3;
            while (dr.Read())
            {
                excel_app.Cells[j, 1].Value = String.Format("{0}", dr["n_mast"]);
                excel_app.Cells[j, 2].Value = String.Format("{0}", dr["fio"]);
                excel_app.Cells[j, 3].Value = String.Format("{0}", dr["dolg"]);
                
                Excel.Range curr_cells = (Excel.Range)excel_app.get_Range("A" + j, "C" + j).Cells;
                curr_cells.Font.Size = 12;
                curr_cells.Borders.LineStyle = 1;

                j = j + 1;
            }
            dr.Close();
            con1.Close();
        }

        private void выполнениеУслугЗаПериодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormUslZaPeriod f1 = new FormUslZaPeriod();
            f1.ShowDialog();

        }

        private void рейтингУслугToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormDiagramRUsl f1 = new FormDiagramRUsl();
            f1.ShowDialog();
        }
    }
}
