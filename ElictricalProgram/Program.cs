using System;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Elictrical_Program
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DBFunctions.conn = new OleDbConnection();
            string database_path = Application.StartupPath + "\\db";
            DBFunctions.conn.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data source=" + database_path + ";Jet OLEDB:Database Password=" + DBFunctions.Base64Decode("RDA1XjczMiMzMEBCQzgqOTA1Q3g=");
            if (!DBFunctions.isAppValid())
            {
                MessageBox.Show("Application is not valid. Exiting...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return; // Ensure that the application does not continue to run
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
