using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace DefectNameReportVersion10
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string _SourceSQLConn = ConfigurationManager.AppSettings["SourceSQLConn"];
            SqlConnection sqlConn = null;
            sqlConn = new SqlConnection(_SourceSQLConn);
            if (sqlConn.State != ConnectionState.Open)
            {
                sqlConn.Close(); sqlConn.Open();
            }
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = sqlConn;
            string comStr = "";
            comStr = string.Format("Select Version from [dbo].[Version_Control] where FileName='AOIVRSDefectNameEXCELReport'");
            cmd.CommandText = comStr;
            //cmd.Parameters.AddWithValue("@lot", lot);
            SqlDataAdapter sa = new SqlDataAdapter(cmd);
            sa = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            sa.Fill(dt);
            if (Equals(dt.Rows[0][0], "Ver1.0.0.0"))
            {
                Application.Run(new Form1());
            }
            else
            {
                MessageBox.Show("The current version is too low, please contact the administrator ( Felix Sun )!");
            }
        }
    }
}
