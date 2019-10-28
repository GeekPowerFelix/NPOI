using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;



namespace DefectNameReportVersion10
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show("Whether you want to close the form?", "Notes:", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr == DialogResult.Yes)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.AcceptButton = button1;    //Enter
            this.CancelButton = button1;    //Esc
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            String lotid = textBox1.Text;
            String Path = comboBox1.Text;
            try
            {
                if (lotid == "" || Path == "")
                {
                    MessageBox.Show("lot id or path is null,reinput pls!");
                }
                //A.create table aoi/vrs
                DataTable dt_aoi = SourceData.GetData.GetAOISourceData(lotid);
                DataTable dt_vrs = SourceData.GetData.GetVRSSourceData(lotid);
                //String filePath = @"D:\ATS\felixsun\Desktop\Test\test.xls";
                //String TPath = Path + "\\" + lotid + "_" + @"AOIVRSDefectNameReport.xls";
                //String aoiPath = Path + "\\" + lotid + "_" + @"AOIDefectNameReport.xls";
                //String vrsPath = Path + "\\" + @lotid + "_" + @"VRSDefectNameReport.xls";
                //String aoiSheetName = "AOI";
                //String vrsSheetName = "VRS";
                //bool IsWriteColumnName = true;

                //if (!File.Exists(aoiPath))
                //{
                //    FileStream fs = File.Create(aoiPath);
                //    fs.Close();
                //    Functions.OperateExcel.ExcelHelper.DataTableToExcel(aoiPath, dt_aoi, aoiSheetName, IsWriteColumnName);
                //    MessageBox.Show("Congratulations, you have successfully created!");
                //}

                DataSet ds = new DataSet();
                ds.Tables.Add(dt_aoi);
                ds.Tables.Add(dt_vrs);

                //String TargetFileNamePath = @"D:\ATS\felixsun\Desktop\Test\test000.xls";
                String TargetFileNamePath = comboBox1.Text +"\\"+"DefectNameReportV1.0.0.xls";
                if (!File.Exists(TargetFileNamePath))
                {
                    FileStream fs = File.Create(TargetFileNamePath);
                    fs.Close();
                }
                else
                {
                    FileStream fs = File.OpenWrite(TargetFileNamePath);
                    fs.Close();
                }

                Functions.OperateExcelMulSheet.DataTableToExcel(ds,TargetFileNamePath);
                MessageBox.Show("Congratulations, you have successfully created!");
                //B.create excel(book,sheet)
                //1.create book
                //创建工作簿对象
                //根据Excel文件的后缀名创建对应的workbook
                //此处固定一个工作簿
                //一定要注意，此处还没有创建excel,若需要创建，代码如下：
                //if (!File.Exists(Path))
                //{
                //    FileStream fs = File.Create(aoiPath);
                //    fs.Close();
                //    //Functions.OperateExcel.ExcelHelper.DataTableToExcel(aoiPath, dt_aoi, aoiSheetName, IsWriteColumnName);
                //    MessageBox.Show("Congratulations, you have successfully created!");
                //}
                //此时在缓存中虚构的excel
                //判断是哪一种的excel 



            }
            catch (Exception ex1)
            {
                throw new Exception(ex1.Message);
            }
            finally
            {
                //KillExcelProcess(excel);//结束Excel进程                
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
    }
}
