using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using INI;
namespace cost_management
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// 
        /// </summary>
        ///
        public static Frmloading frmload;
        public static frmDisplayReports frmdata;
        [STAThread]
        static void Main()
        {
            frmload = new Frmloading();
            frmdata = new frmDisplayReports();
            ClsCommonFunction clsFun = new ClsCommonFunction();

            try
            {
                if (!clsFun.Openconnection().State.Equals(System.Data.ConnectionState.Open))
                {
                    Application.Exit();
                    return;
                }
                
                Application.Run(new MidHome());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                Application.Exit();
            }

        }
    }
}
