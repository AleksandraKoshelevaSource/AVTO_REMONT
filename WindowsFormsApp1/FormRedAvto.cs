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

namespace WindowsFormsApp1
{
    public partial class FormRedAvto : System.Windows.Forms.Form
    {
        string navto = "1";
        public FormRedAvto()
        {
            InitializeComponent();
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
            dataGridView1.Columns[1].Width = 160;
            dataGridView1.Columns[2].Width = 90;
            dataGridView1.Columns[3].Width = 90;
            dataGridView1.Columns[4].Width = 60;
            dataGridView1.Columns[5].Width = 60;
            dataGridView1.Columns[6].Width = 100;
            //dataGridView1.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 10);
        }
        //
        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(Data.Glob_connection_string);
            con.Open();

            string SQL_text = "SELECT * FROM AVTO WHERE vin LIKE N'%" + textBox1.Text + "%'";
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt;
            ChangeGrid();
            con.Close();
        
        }

        private void FormRedAvto_Activated(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(Data.Glob_connection_string);
            con.Open();

            string SQL_text = "SELECT * FROM AVTO";
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt;
            ChangeGrid();
            con.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection(Data.Glob_connection_string);
            con.Open();

            string SQL_text = "SELECT * FROM AVTO WHERE fio_v LIKE N'%" + textBox2.Text + "%'";
            SqlDataAdapter da = new SqlDataAdapter(SQL_text, con);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.Refresh();
            dataGridView1.DataSource = dt;
            ChangeGrid();
            con.Close();
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            textBox3.Text = dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString();
            textBox4.Text = dataGridView1[2, dataGridView1.CurrentRow.Index].Value.ToString();
            textBox5.Text = dataGridView1[3, dataGridView1.CurrentRow.Index].Value.ToString();
            textBox6.Text = dataGridView1[4, dataGridView1.CurrentRow.Index].Value.ToString();
            textBox7.Text = dataGridView1[5, dataGridView1.CurrentRow.Index].Value.ToString();
            textBox8.Text = dataGridView1[6, dataGridView1.CurrentRow.Index].Value.ToString();
            navto = dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string SQL_text = "UPDATE AVTO SET vin = N'" + textBox3.Text + "', marka = N'" + textBox4.Text + "', model = N'" +
                textBox5.Text + "', god = N'" + textBox6.Text + "', reg_n = N'" + textBox7.Text +
                "', fio_v = N'" + textBox8.Text +
                "' WHERE n_avto = " + navto;
            SqlConnection con = new SqlConnection(Data.Glob_connection_string);
            con.Open();
            SqlCommand com1 = new SqlCommand(SQL_text, con);
            SqlDataReader dr = com1.ExecuteReader();
            
            //MessageBox.Show("Данные изменены");
            this.Activate();
            dr.Close();
            con.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
             string SQL_text = "DELETE FROM AVTO WHERE n_avto = " + navto;
            SqlConnection con = new SqlConnection(Data.Glob_connection_string);
            con.Open();
            SqlCommand com1 = new SqlCommand(SQL_text, con);
            SqlDataReader dr = com1.ExecuteReader();
            dr.Close();
            con.Close();
            //MessageBox.Show("Данные удалены");
        }
    }
}
