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
    public partial class FormUsl : System.Windows.Forms.Form
    {
        //string connectionString = @"Data Source = (LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Home-PC\source\repos\Remont\WindowsFormsApp1\AVTO_BASE.mdf;Integrated Security = True; Connect Timeout = 30";
        //string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB; AttachDbFilename = |DataDirectory|\AVTO_BASE.mdf; Integrated Security = True";
        string n_usl;
        public FormUsl()
        {
            InitializeComponent();
        }


        private void FormUsl_Activated(object sender, EventArgs e)
        {
            SqlConnection connection1 = new SqlConnection(Data.Glob_connection_string);
            connection1.Open();

            string SQL_select = "SELECT * FROM USLUGI";

            SqlDataAdapter adapter = new SqlDataAdapter(SQL_select, connection1);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            dataGridView1.Refresh();
            dataGridView1.DataSource = dt;
            dataGridView1.Columns[0].HeaderText = "№ услуги";
            dataGridView1.Columns[1].HeaderText = "Наименование";
            dataGridView1.Columns[2].HeaderText = "Стоимость";

            dataGridView1.Columns[0].Width = 100;
            dataGridView1.Columns[1].Width = 300;
            dataGridView1.Columns[2].Width = 200;

            button2.Enabled = false;
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            textBox1.Text = dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString();
            textBox2.Text = dataGridView1[2, dataGridView1.CurrentRow.Index].Value.ToString();
            n_usl = dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString();
            button1.Enabled = false;
            button2.Enabled = true;
        }

        private string change_comma(string s)
        {
            int pos = s.IndexOf(",");
            if (pos > 0)
            {
                s = s.Substring(0, pos) + "." + s.Substring(pos + 1, 2);
            }
            return s;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            string SQL_text = "SELECT max(n_usl) as max_n FROM USLUGI";
            SqlCommand comm1 = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = comm1.ExecuteReader();
            int max_n = 1;
            while (dr.Read())
            {
                max_n = Convert.ToInt32(dr["max_n"]);
            }
            dr.Close();
     
            SQL_text = "INSERT INTO USLUGI(n_usl, naimen, stoim) VALUES (" + (max_n + 1) + 
                ", N'" + textBox1.Text + "', " + change_comma(textBox2.Text) + ")";
            
            comm1 = new SqlCommand(SQL_text, con1);
            dr = comm1.ExecuteReader();
            dr.Close();
            con1.Close();
            MessageBox.Show("Данные сохранены");
            this.Activate();
            this.button3_Click(this, new EventArgs());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //update
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            
            string SQL_text = "UPDATE USLUGI SET naimen = N'" + textBox1.Text + "', stoim = " + 
                change_comma(textBox2.Text) + " WHERE n_usl = " + n_usl;
           
            SqlCommand comm1 = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = comm1.ExecuteReader();
            dr.Close();
            con1.Close();
            MessageBox.Show("Данные изменены");
            this.Activate();
        }

        private void FormUsl_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            button1.Enabled = true;
            button2.Enabled = false;
        }
    }
}
