using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;

namespace DefectNameReportVersion10.Functions
{
    class CreatTable
    {
        public static DataTable CreateAOIOriginalTable()
        {
            DataTable dt1 = new DataTable("AOI Original table");
            dt1.Columns.Add("DateTime", typeof(DateTime));
            dt1.Columns.Add("Machine", typeof(String));
            dt1.Columns.Add("Operator", typeof(String));
            dt1.Columns.Add("Job name", typeof(String));
            dt1.Columns.Add("Layer", typeof(String));
            dt1.Columns.Add("Lot", typeof(String));
            dt1.Columns.Add("2DID", typeof(String));
            dt1.Columns.Add("PanelID", typeof(String));
            dt1.Columns.Add("Defect Total", typeof(Int32));
            dt1.Columns.Add("Pass Total", typeof(Int32));
            dt1.Columns[0].DefaultValue = DateTime.Now;
            dt1.Columns[1].DefaultValue = "";
            dt1.Columns[2].DefaultValue = "";
            dt1.Columns[3].DefaultValue = "";
            dt1.Columns[4].DefaultValue = "";
            dt1.Columns[5].DefaultValue = "";
            dt1.Columns[6].DefaultValue = "";
            dt1.Columns[7].DefaultValue = "";
            dt1.Columns[8].DefaultValue = 0;//for cycle only default defect column
            dt1.Columns[9].DefaultValue = 0;

            //fill  column  according to defect name from db MES_DefectToBinCode
            string _SourceSQLConn = ConfigurationManager.AppSettings["SourceSQLConn"];
            SqlConnection sqlConn = null;
            try
            {
                //dt1-aoi fill in  column defect from dt11
                sqlConn = new SqlConnection(_SourceSQLConn);
                if (sqlConn.State != ConnectionState.Open)
                {
                    sqlConn.Close(); sqlConn.Open();
                }
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlConn;
                string comStr = "";
                comStr = string.Format("Select Defect_Name from MES_DefectToBinCode where Station='AOI' order by Defect_Name");
                cmd.CommandText = comStr;
                DataTable dt2 = new DataTable("AOIDistinct2did");
                SqlDataAdapter sa = new SqlDataAdapter(cmd);
                sa = new SqlDataAdapter(cmd);
                sa.Fill(dt2);
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    dt1.Columns.Add(dt2.Rows[i][0].ToString(), typeof(Int32));
                    dt1.Columns[dt2.Rows[i][0].ToString()].DefaultValue = 0;    //default value
                }
                return dt1;
            }
            catch (Exception ex1)
            {
                throw new Exception(ex1.Message);
            }
            finally
            {
                if (sqlConn != null)
                {
                    sqlConn.Close();
                }
            }
        }
        public static DataTable CreateVRSOriginalTable()
        {
            //DataTable dt = CreateAOIOriginalTable();
            //DataTable dt1 = new DataTable();
            //dt1 = dt.Clone();     //error defect column is different
            DataTable dt1 = new DataTable("VRS Original table");
            dt1.Columns.Add("DateTime", typeof(DateTime));
            dt1.Columns.Add("Machine", typeof(String));
            dt1.Columns.Add("Operator", typeof(String));
            dt1.Columns.Add("Job name", typeof(String));
            dt1.Columns.Add("Layer", typeof(String));
            dt1.Columns.Add("Lot", typeof(String));
            dt1.Columns.Add("2DID", typeof(String));
            dt1.Columns.Add("PanelID", typeof(String));
            dt1.Columns.Add("Defect Total", typeof(Int32));
            dt1.Columns.Add("Pass Total", typeof(Int32));
            dt1.Columns[0].DefaultValue = DateTime.Now;
            dt1.Columns[1].DefaultValue = "";
            dt1.Columns[2].DefaultValue = "";
            dt1.Columns[3].DefaultValue = "";
            dt1.Columns[4].DefaultValue = "";
            dt1.Columns[5].DefaultValue = "";
            dt1.Columns[6].DefaultValue = "";
            dt1.Columns[7].DefaultValue = "";
            dt1.Columns[8].DefaultValue = 0;//for cycle only default defect column
            dt1.Columns[9].DefaultValue = 0;
            string _SourceSQLConn = ConfigurationManager.AppSettings["SourceSQLConn"];
            SqlConnection sqlConn = null;
            try
            {
                sqlConn = new SqlConnection(_SourceSQLConn);
                if (sqlConn.State != ConnectionState.Open)
                {
                    sqlConn.Close(); sqlConn.Open();
                }
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = sqlConn;
                string comStr = "";
                comStr = string.Format("Select Defect_Name from MES_DefectToBinCode where Station='VRS' order by Defect_Name");
                cmd.CommandText = comStr;
                DataTable dt2 = new DataTable();
                SqlDataAdapter sa = new SqlDataAdapter(cmd);
                sa = new SqlDataAdapter(cmd);
                sa.Fill(dt2);
                //dt1-vrs fill in  defect column from dt2
                for (int i = 0; i < dt2.Rows.Count; i++)
                {
                    dt1.Columns.Add(dt2.Rows[i][0].ToString(), typeof(Int32));
                    dt1.Columns[dt2.Rows[i][0].ToString()].DefaultValue = 0;
                }
                return dt1;
            }
            catch (Exception ex1)
            {
                throw new Exception(ex1.Message);
            }
            finally
            {
                if (sqlConn != null)
                {
                    sqlConn.Close();
                }
            }

        }
    }
}
