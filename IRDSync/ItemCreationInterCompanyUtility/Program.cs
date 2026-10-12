using IRDSync;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace IRDSync
{
    static class Program
    {
        //public static IRDSyncDairy Dxt;
        ///// <summary>
        ///// The main entry point for the application.
        ///// </summary>
        //[STAThread]
        //static void Main()
        //{

        //    Application.EnableVisualStyles();
        //    Application.SetCompatibleTextRenderingDefault(false);
        //    Dxt= new IRDSyncDairy();
        //    Application.Exit();
        //   // Application.Run(new FrmCompanySelection());

        //}

        //static void Main(string[] args)
        //{
        //    try
        //    {
        //        IRDSyncDairy logic = new IRDSyncDairy();
        //        logic.ProcessSync();
        //    }
        //    catch (Exception ex)
        //    {
        //        //Console.WriteLine("Fatal: " + ex.Message);
        //    }
        //}

        static void Main(string[] args)
        {
            IRDSyncDairy sync = new IRDSyncDairy();
            sync.StartProcess();
        }
    }
}
