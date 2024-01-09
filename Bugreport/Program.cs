using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bugreport
{
    internal static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            StartScreen startScreen = new StartScreen();
            DateTime end = DateTime.Now + TimeSpan.FromSeconds(2);
            startScreen.Show();
            while (end > DateTime.Now)
            {
                Application.DoEvents();
            }
            startScreen.Close();
            startScreen.Dispose();
            Application.Run(new Form1());
        }
    }
}
