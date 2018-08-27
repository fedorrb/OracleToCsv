using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace LsDumpToCsv
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            string arg0 = "";
            if (args.Length == 0)
                arg0 = "none";
            else
                arg0 = args[0];
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(arg0));
        }
    }
}
