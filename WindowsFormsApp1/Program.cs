using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    static class Data
    {
        public static string Glob_connection_string  = @"Data Source = (LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\Home-PC\source\repos\Remont\WindowsFormsApp1\AVTO_BASE.mdf;Integrated Security = True; Connect Timeout = 30";
        
    }
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new F_Menu());
        }
    }
}
