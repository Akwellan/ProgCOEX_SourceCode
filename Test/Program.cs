using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test
{
    class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            //string EtatTimer = "18/11/2020 12:29";
            //string Date = EtatTimer.Substring(0, EtatTimer.Length - 6);
            //string Heure = EtatTimer.Substring(11);
            //Console.WriteLine("Date :" + Date + "Heure :" + Heure);
        }
    }
}
