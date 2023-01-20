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
    public partial class FormDobAvto : System.Windows.Forms.Form
    {
        //string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Home-PC\source\repos\Remont\WindowsFormsApp1\AVTO_BASE.mdf;Integrated Security=True;Connect Timeout=30";
        //string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\AVTO_BASE.mdf;Integrated Security=True;Connect Timeout=30";
        public FormDobAvto()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();

            int n_avto = 1;
            string SQL_text = "SELECT max(n_avto) as max_n FROM AVTO";
            SqlCommand cm1 = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = cm1.ExecuteReader();
            while (dr.Read())
            {
                string sn_avto = String.Format("{0}", dr["max_n"]);
                if (sn_avto != "")
                {
                    n_avto = Convert.ToInt32(sn_avto);
                    n_avto++;
                }
                
            }
            dr.Close();
            
            //INSERT INTO AVTO VALUES(nnn, N'vvv', N'mmm', N'mmm', N'YYYY', N'rrr', N'fff')
            SQL_text = "INSERT INTO AVTO(n_avto, vin, marka, model, god, reg_n, fio_v) VALUES(" + 
                n_avto + ", N'" + textBox1.Text + "', N'" + textBox2.Text + 
                "', N'" + textBox3.Text + "', N'" + textBox4.Text + "', N'" + 
                textBox5.Text +"', N'" + textBox6.Text + "')";
            //MessageBox.Show(SQL_text);
            cm1 = new SqlCommand(SQL_text, con1);
            dr = cm1.ExecuteReader();
            dr.Close();
            con1.Close();
            MessageBox.Show("Данные сохранены");
            this.Close();
        }
    }
}
