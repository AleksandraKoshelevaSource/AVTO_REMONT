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
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp1
{
    public partial class FormRemontAvto : System.Windows.Forms.Form
    {
        private Excel.Application excel_app;
        private double summaUslugi = 0;
        private double stoimUslugi = 0;
        
        private int i = 0;
        //string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Home-PC\source\repos\Remont\WindowsFormsApp1\AVTO_BASE.mdf;Integrated Security = True; Connect Timeout = 30";
        //string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB; AttachDbFilename = |DataDirectory|\AVTO_BASE.mdf; Integrated Security = True";
        public FormRemontAvto()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FormDobAvto fd = new FormDobAvto();
            fd.ShowDialog();
        }

        private void RaschetSummi()
        {
            summaUslugi = (stoimUslugi * Convert.ToDouble(textBox3.Text));
        }

        private void ChangeGrid()
        {
            dataGridView1.Columns[0].HeaderText = "№";
            dataGridView1.Columns[1].HeaderText = "VIN";
            dataGridView1.Columns[2].HeaderText = "Марка";
            dataGridView1.Columns[3].HeaderText = "Модель";
            dataGridView1.Columns[4].HeaderText = "Год";
            dataGridView1.Columns[5].HeaderText = "Рег. №";
            dataGridView1.Columns[6].HeaderText = "ФИО владельца";
            dataGridView1.Columns[0].Width = 40;
            dataGridView1.Columns[1].Width = 190;
            dataGridView1.Columns[2].Width = 120;
            dataGridView1.Columns[3].Width = 120;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].Width = 100;
            dataGridView1.Columns[6].Width = 150;
        }

        private void FormRemontAvto_Activated(object sender, EventArgs e)
        {
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            string SQL_text = "SELECT * FROM AVTO";
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con1);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt;
            con1.Close();
            ChangeGrid();

            SQL_text = "SELECT n_mast, fio, dolg FROM MASTER";
            con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            da = new SqlDataAdapter(SQL_text,con1);
            dt = new DataTable();
            da.Fill(dt);
            comboBox1.DataSource = dt;
            comboBox1.DisplayMember = "fio";
            comboBox1.ValueMember = "n_mast";
            con1.Close();

            SQL_text = "SELECT n_usl, naimen, stoim FROM USLUGI";
            con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            da = new SqlDataAdapter(SQL_text, con1);
            dt = new DataTable();
            da.Fill(dt);
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "naimen";
            comboBox2.ValueMember = "n_usl";
            con1.Close();

            textBox3.Text = "1";

            Otobr_stoim();
            Otobr_uslugi();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string SQL_text = "SELECT * FROM AVTO WHERE vin LIKE N'" + textBox1.Text + "%'";
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con1);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt;
            ChangeGrid();
            con1.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string SQL_text = "SELECT * FROM AVTO WHERE fio_v LIKE N'" + textBox2.Text + "%'";
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con1);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt;
            ChangeGrid();
            con1.Close();
        }

        private void New_zakaz_naryad()
        {
            string SQL_text = "SELECT max(n_z_n) as max_z_n FROM REMONT";
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            SqlCommand com1 = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = com1.ExecuteReader();

            int n_z_n = 1;
            while (dr.Read())
            {
                string sn_z_n = String.Format("{0}", dr["max_z_n"]);
                if (sn_z_n != "")
                {
                    n_z_n = Convert.ToInt32(sn_z_n);
                    n_z_n++;
                }
            }
            dr.Close();
            con1.Close();
            label6.Text = String.Format("{0}", n_z_n);
        }

        private void FormRemontAvto_Load(object sender, EventArgs e)
        {

            New_zakaz_naryad();
        }

        private string change_sum(string stoim, string kol)
        {
            double summa = (Convert.ToDouble(stoim) * Convert.ToDouble(kol));
            return Convert.ToString(summa);
        }

        private void Otobr_stoim()
        {
            textBox3.Text = "1";
            //otobr stoimost
            string SQL_text = "SELECT stoim FROM USLUGI WHERE n_usl=" + comboBox2.SelectedValue;
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            SqlCommand com1 = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = com1.ExecuteReader();
            while (dr.Read())
            {
                stoimUslugi = Convert.ToDouble(dr["stoim"]);
            }
            RaschetSummi();
            dr.Close();
            con1.Close();
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            textBox3.Text = "1";
            //otobr stoimost
            Otobr_stoim();
        }

        private void textBox3_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                RaschetSummi();
            }
        }

        private void Otobr_remont()
        {
            string SQL_text = "SELECT R.id, M.fio, U.naimen, U.stoim, R.kol, U.stoim * R.kol " +
                "FROM REMONT R, USLUGI U, MASTER M " +
                "WHERE R.n_usl = U.n_usl AND R.n_mast = M.n_mast " +
                "AND R.n_z_n = " + label6.Text;
            //MessageBox.Show(SQL_text);

            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con1);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView2.Refresh();
            dataGridView2.DataSource = dt;

            dataGridView2.Columns[0].HeaderText = "№";
            dataGridView2.Columns[1].HeaderText = "Мастер";
            dataGridView2.Columns[2].HeaderText = "Услуга";
            dataGridView2.Columns[3].HeaderText = "Цена";
            dataGridView2.Columns[4].HeaderText = "Количество";
            dataGridView2.Columns[5].HeaderText = "Сумма";

            dataGridView2.Columns[0].Width = 25;
            dataGridView2.Columns[1].Width = 100;
            dataGridView2.Columns[2].Width = 120;
            dataGridView2.Columns[3].Width = 60;
            dataGridView2.Columns[4].Width = 60;
            dataGridView2.Columns[5].Width = 60;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Вычислить id 
            string SQL_text = "SELECT max(id) as max_id FROM REMONT";
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            SqlCommand com1 = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = com1.ExecuteReader();
            int new_id = 1;
            while (dr.Read())
            {
                string s_max_id = String.Format("{0}",dr["max_id"]);
                if (s_max_id != "")
                {
                    new_id = Convert.ToInt32(s_max_id);
                    new_id++;
                }
            }
            dr.Close();
            con1.Close();

            //посчитать сумму услуги = стоимость * количество
            RaschetSummi();

            //сохранить данные в таблицу REMONT
            string n_avto = dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString();

            SQL_text = "INSERT INTO REMONT(id, n_avto,n_usl, n_mast, data, kol, n_z_n, sum) VALUES (" +
                new_id + ", " + n_avto + ", " + comboBox2.SelectedValue + ", " + comboBox1.SelectedValue + ", '" +
                dateTimePicker1.Value.ToString("MM/dd/yyyy") + "', " + textBox3.Text + ", " + label6.Text + ", " + summaUslugi + ")";
            //MessageBox.Show(SQL_text);
            con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            com1 = new SqlCommand(SQL_text, con1);
            dr = com1.ExecuteReader();
            dr.Close();
            con1.Close();
            //MessageBox.Show("Данные сохранены");

            //Отобразить данные в гриде
            Otobr_remont();

            con1.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string id = dataGridView2[0, dataGridView2.CurrentRow.Index].Value.ToString();

            string SQL_text = "DELETE FROM REMONT WHERE id = " + id;
            SqlConnection con = new SqlConnection(Data.Glob_connection_string);
            con.Open();

            SqlCommand cmd = new SqlCommand(SQL_text, con);
            SqlDataReader dr = cmd.ExecuteReader();
            dr.Close();
            con.Close();

            Otobr_remont();
        }

        private void Otobr_uslugi()
        {
            string SQL_text = "SELECT R.n_z_n, R.data, sum(sum) FROM REMONT R GROUP BY R.n_z_n, R.data";
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con1);
            DataTable dt = new DataTable();
            da.Fill(dt);


            dataGridView3.Refresh();
            dataGridView3.DataSource = dt;

            dataGridView3.Columns[0].HeaderText = "№";
            dataGridView3.Columns[1].HeaderText = "Дата";
            dataGridView3.Columns[2].HeaderText = "Сумма";

            dataGridView3.Columns[0].Width = 40;
            dataGridView3.Columns[1].Width = 120;
            dataGridView3.Columns[2].Width = 120;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            New_zakaz_naryad();
            Otobr_remont();
            Otobr_uslugi();
        }

        private void dataGridView3_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            excel_app = new Excel.Application();
            excel_app.Visible = true;
            excel_app.SheetsInNewWorkbook = 1;
            excel_app.Workbooks.Add(Type.Missing);

            Excel.Range _excelCells = (Excel.Range)excel_app.get_Range("A1", "F1").Cells;
            _excelCells.Merge(Type.Missing);
            _excelCells = (Excel.Range)excel_app.get_Range("A2", "F2").Cells;
            _excelCells.Merge(Type.Missing);
            _excelCells = (Excel.Range)excel_app.get_Range("A3", "F3").Cells;
            _excelCells.Merge(Type.Missing);
            _excelCells = (Excel.Range)excel_app.get_Range("A4", "F4").Cells;
            _excelCells.Merge(Type.Missing);
            _excelCells = (Excel.Range)excel_app.get_Range("A5", "F5").Cells;
            _excelCells.Merge(Type.Missing);
            _excelCells = (Excel.Range)excel_app.get_Range("A6", "F6").Cells;
            _excelCells.Merge(Type.Missing);
            _excelCells = (Excel.Range)excel_app.get_Range("A7", "F7").Cells;
            _excelCells.Merge(Type.Missing);

            excel_app.Cells[1, 1].Value = "Заказ-наряд № " + label6.Text + " от " + dateTimePicker1.Value.ToString("MM/dd/yyyy");
            excel_app.Cells[1, 1].Font.Bold = true;
            excel_app.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            excel_app.Cells[2, 1].Value = "Исполнитель: Техцентр AVTO, ИНН 26311001412544, тел.: 7-77-77";
            excel_app.Cells[3, 1].Value = "Заказчик: " + dataGridView1[6, dataGridView1.CurrentRow.Index].Value.ToString();
            excel_app.Cells[4, 1].Value = "Модель: " + dataGridView1[3, dataGridView1.CurrentRow.Index].Value.ToString() +
                 " " + dataGridView1[2, dataGridView1.CurrentRow.Index].Value.ToString();
            excel_app.Cells[5, 1].Value = "Год выпуска: " + dataGridView1[4, dataGridView1.CurrentRow.Index].Value.ToString();
            excel_app.Cells[6, 1].Value = "Рег. №: " + dataGridView1[5, dataGridView1.CurrentRow.Index].Value.ToString();
            excel_app.Cells[7, 1].Value = "Год выпуска: " + dataGridView1[4, dataGridView1.CurrentRow.Index].Value.ToString();
           

            for (int i = 1; i <= 7; i++)
            {
                excel_app.Cells[i, 1].Font.Size = 14;
                excel_app.Cells[i, 1].Font.Italic = true;
            }

                excel_app.Cells[9,1].Value = "№";
            excel_app.Columns[1].columnwidth = 3;

            excel_app.Cells[9, 2].Value = "Наименование работы";
            excel_app.Columns[2].columnwidth = 30;

            excel_app.Cells[9, 3].Value = "ФИО работника";
            excel_app.Columns[3].columnwidth = 30;

            excel_app.Cells[9, 4].Value = "Цена";
            excel_app.Columns[4].columnwidth = 12;

            excel_app.Cells[9, 5].Value = "Кол-во";
            excel_app.Columns[5].columnwidth = 12;

            excel_app.Cells[9, 6].Value = "Сумма";
            excel_app.Columns[6].columnwidth = 12;

            for (int i = 1; i<=6; i++)
            {
                excel_app.Cells[9, i].Font.Size = 14;
                excel_app.Cells[9, i].Font.Italic = true;
                excel_app.Cells[9, i].Font.Bold = true;
                excel_app.Cells[9, i].Borders.LineStyle = 1;
                excel_app.Cells[9, i].Borders.Weight = Excel.XlBorderWeight.xlThick;
            }
           

            string SQL_text = "SELECT R.id, M.fio, U.naimen, U.stoim, R.kol, U.stoim * R.kol as summa " + 
                "FROM REMONT R, USLUGI U, MASTER M WHERE " + 
                "R.n_usl = U.n_usl AND M.n_mast = R.n_mast AND R.n_z_n = " + label6.Text;
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            SqlCommand comm = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = comm.ExecuteReader();
            i = 10;
            decimal itog_summa = 0; 
            while (dr.Read())
            {
                excel_app.Cells[i,1].Value = i - 9;
                excel_app.Cells[i, 2].Value = String.Format("{0}", dr["naimen"]);
                excel_app.Cells[i, 3].Value = String.Format("{0}", dr["fio"]);
                excel_app.Cells[i, 4].Value = String.Format("{0}", dr["stoim"]);
                excel_app.Cells[i, 5].Value = String.Format("{0}", dr["kol"]);
                excel_app.Cells[i, 6].Value = String.Format("{0}", dr["summa"]);

                Excel.Range curr_cells = (Excel.Range)excel_app.get_Range("A" + i, "F" + i).Cells;
                curr_cells.Font.Size = 12;
                curr_cells.Borders.LineStyle = 1;

                itog_summa = itog_summa + Convert.ToDecimal(dr["summa"]);
                i = i + 1;
            }
            dr.Close();
            con1.Close();
            excel_app.Cells[i, 5].Value = "ИТОГО";
            excel_app.Cells[i, 5].Font.Size = 12;
            excel_app.Cells[i, 5].Borders.LineStyle = 1;
            excel_app.Cells[i, 6].Value = itog_summa;
            excel_app.Cells[i, 6].Font.Size = 12;
            excel_app.Cells[i, 6].Borders.LineStyle = 1;

            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
