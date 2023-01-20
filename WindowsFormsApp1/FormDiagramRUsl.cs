using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data.SqlClient;

namespace WindowsFormsApp1
{
    public partial class FormDiagramRUsl : Form
    {
        public FormDiagramRUsl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart1.Titles.Clear();
            chart1.Series.RemoveAt(0);
            chart1.Palette = ChartColorPalette.SeaGreen;
            string diagTitle = "Рейтинг услуг";
            chart1.Titles.Add(diagTitle);
            Series s1 = new Series("Услуги");
            s1.Color = Color.OrangeRed;
            string SQL_text = "SELECT u.naimen, u.stoim, sum(r.kol) as kol FROM USLUGI u, REMONT r WHERE u.n_usl=r.n_usl " + 
                " AND r.data >= '" + dateTimePicker1.Value.ToString("yyyyMMdd") + 
                "' AND r.data <= '" + dateTimePicker2.Value.ToString("yyyyMMdd") + 
                "' GROUP BY u.naimen, u.stoim";
            SqlConnection con1 = new SqlConnection(Data.Glob_connection_string);
            con1.Open();
            SqlCommand com1 = new SqlCommand(SQL_text, con1);
            SqlDataReader dr = com1.ExecuteReader();
            string naim = "";
            int kol = 0;
            while (dr.Read())
            {
                naim = Convert.ToString(dr["naimen"]);
                kol = Convert.ToInt32(dr["kol"]);
                s1.Points.AddXY(naim, kol);
            }
            dr.Close();
            con1.Close();
            chart1.Series.Add(s1);
        }
    }
}
