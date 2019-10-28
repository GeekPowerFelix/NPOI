using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace DefectNameReportVersion10.SourceData
{
    class GetData
    {
        //fill 2did then fill rest data
        public static DataTable GetAOISourceData(string lot)
        {
            DataTable dt1 = Functions.CreatTable.CreateAOIOriginalTable();
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
                comStr = string.Format("select distinct MAX(DateTesting) as LastestTime, Machine_Name,Operator_Name,Product_ID,Layer_Name,LotNO_MES,BC_2DID,Panel_ID,Defect_Name,COUNT(Defect_Name) as DefectNameNumber from Orbo_AOI_Record where LotNO_MES=@lot and Defect_Name is not null Group by Machine_Name,Operator_Name,Product_ID,LotNO_MES, Layer_Name,BC_2DID, Panel_ID, Defect_Name Order by LastestTime");
                cmd.CommandText = comStr;
                cmd.Parameters.AddWithValue("@lot", lot);
                //cmd.Parameters.AddWithValue("@lot", "613000000598");
                DataTable dt = new DataTable();
                SqlDataAdapter sa = new SqlDataAdapter(cmd);
                sa = new SqlDataAdapter(cmd);
                sa.Fill(dt);
                DataView dataView = dt.DefaultView;
                DataTable dataTableDistinct = dataView.ToTable(true, "BC_2DID");
                //fill in 2did;
                for (int i = 0; i < dataTableDistinct.Rows.Count; i++)
                {
                    dt1.Rows.Add();
                    dt1.Rows[i][6] = dataTableDistinct.Rows[i][0].ToString();
                }
                //fill all data except total
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt1.Rows.Count; j++)
                    {
                        if (dt1.Rows[j][6].ToString() == dt.Rows[i][6].ToString())
                        {
                            dt1.Rows[j][0] = dt.Rows[i][0];
                            dt1.Rows[j][1] = dt.Rows[i][1];
                            dt1.Rows[j][2] = dt.Rows[i][2];
                            dt1.Rows[j][3] = dt.Rows[i][3];
                            dt1.Rows[j][4] = dt.Rows[i][4];
                            dt1.Rows[j][5] = dt.Rows[i][5];
                            //dt1.Rows[j][6] = dt.Rows[i][6];
                            dt1.Rows[j][7] = dt.Rows[i][7];

                            dt1.Rows[j][dt.Rows[i][8].ToString()] = int.Parse(dt.Rows[i][9].ToString());
                        }
                        else
                        {
                        }
                    }
                }

                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    for (int j = 8; j < dt1.Columns.Count; j++)
                    {
                        Int32 a = int.Parse(dt1.Rows[i][8].ToString());
                        Int32 b = int.Parse(dt1.Rows[i][j].ToString());

                        dt1.Rows[i]["Defect Total"] = a + b;
                    }
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
        public static DataTable GetVRSSourceData(string lot)
        {
            DataTable dt1 = Functions.CreatTable.CreateVRSOriginalTable();
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
                comStr = string.Format("select distinct MAX(DateTesting) as LastestTime, Machine_Name,Operator_Name,Product_ID,Layer_Name,LotNO_MES,BC_2DID,Panel_ID,Defect_Name,COUNT(Defect_Name) as DefectNameNumber from Orbo_VRS_Data where LotNO_MES=@lot and Defect_Name is not null Group by Machine_Name,Operator_Name,Product_ID,LotNO_MES, Layer_Name,BC_2DID, Panel_ID, Defect_Name Order by LastestTime");
                cmd.CommandText = comStr;
                cmd.Parameters.AddWithValue("@lot", lot);
                //cmd.Parameters.AddWithValue("@lot", "613000000598");
                DataTable dt = new DataTable();
                SqlDataAdapter sa = new SqlDataAdapter(cmd);
                sa = new SqlDataAdapter(cmd);
                sa.Fill(dt);
                DataView dataView = dt.DefaultView;
                DataTable dataTableDistinct = dataView.ToTable(true, "BC_2DID");
                //fill in 2did;
                for (int i = 0; i < dataTableDistinct.Rows.Count; i++)
                {
                    dt1.Rows.Add();
                    dt1.Rows[i][6] = dataTableDistinct.Rows[i][0];
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt1.Rows.Count; j++)
                    {
                        if (dt1.Rows[j][6].ToString() == dt.Rows[i][6].ToString())
                        {
                            dt1.Rows[j][0] = dt.Rows[i][0];
                            dt1.Rows[j][1] = dt.Rows[i][1];
                            dt1.Rows[j][2] = dt.Rows[i][2];
                            dt1.Rows[j][3] = dt.Rows[i][3];
                            dt1.Rows[j][4] = dt.Rows[i][4];
                            dt1.Rows[j][5] = dt.Rows[i][5];
                            //dt1.Rows[j][6] = dt.Rows[i][6];//
                            dt1.Rows[j][7] = dt.Rows[i][7];

                            dt1.Rows[j][dt.Rows[i][8].ToString()] = int.Parse(dt.Rows[i][9].ToString());
                        }
                        else
                        {
                        }
                    }
                }
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    for (int j = 8; j < dt1.Columns.Count; j++)
                    {
                        if (dt1.Columns[j].ToString() == "False Alarm" || dt1.Columns[j].ToString()== "Code 33"|| dt1.Columns[j].ToString() == "Code 99")
                        {
                            Int32 a = int.Parse(dt1.Rows[i][9].ToString());
                            Int32 b = int.Parse(dt1.Rows[i][j].ToString());
                            dt1.Rows[i]["Pass Total"] = a + b;
                        }
                        else
                        {
                            Int32 a = int.Parse(dt1.Rows[i][8].ToString());
                            Int32 b = int.Parse(dt1.Rows[i][j].ToString());
                            dt1.Rows[i]["Defect Total"] = a + b;
                        }

                    }

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
