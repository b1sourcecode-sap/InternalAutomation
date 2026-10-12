using PDFCreation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PDFCreation
{
    static class Program
    {
        //public static PDFCreationDairy Dxt;
        ///// <summary>
        ///// The main entry point for the application.
        ///// </summary>
        //[STAThread]
        //static void Main()
        //{

        //    Application.EnableVisualStyles();
        //    Application.SetCompatibleTextRenderingDefault(false);
        //    Dxt= new PDFCreationDairy();
        //    Application.Exit();
        //   // Application.Run(new FrmCompanySelection());

        //}

        //static void Main(string[] args)
        //{
        //    try
        //    {
        //        PDFCreationDairy logic = new PDFCreationDairy();
        //        logic.ProcessSync();
        //    }
        //    catch (Exception ex)
        //    {
        //        //Console.WriteLine("Fatal: " + ex.Message);
        //    }
        //}

        static void Main(string[] args)
        {
            PDFCreationDairy sync = new PDFCreationDairy();
            sync.StartProcess();
        }
    }
}
