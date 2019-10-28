using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Threading.Tasks;

namespace DefectNameReportVersion10.Functions
{
    class OperateExcelMulSheet
    {
        public static void DataTableToExcel(DataSet ds,String TargetFileNamePath)
        {
            IWorkbook workbook = new HSSFWorkbook();
            //IWorkbook workbook = null;  //初始化
            if (TargetFileNamePath.IndexOf(".xlsx") > 0)
            {
                workbook = new XSSFWorkbook();  //2007版本的excel
            }
            else if (TargetFileNamePath.IndexOf(".xls") > 0) //2003版本的excel
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                //return;    //都不匹配或者传入的文件根本就不是excel文件，直接返回
            }
            try
            {
                //创建工作表，根据实际需求创建多个sheet
                for (int i = 0; i < ds.Tables.Count; i++)//ds.Tables.Count ->sheet count
                {
                    //获取一个工作表
                    System.Data.DataTable table = ds.Tables[i];
                    //??此处应该加一个对原始数据表名的检查
                    //创建sheet并命名
                    ISheet sheet = workbook.CreateSheet(ds.Tables[i].TableName);
                    //creat rows and columns data
                    //创建Row中的列Cell并赋值【SetCellValue有5个重载方法 bool、DateTime、double、string、IRichTextString(未演示)
                    for (int j = 0; j < table.Rows.Count+1; j++)
                    {
                        //dt1.Rows.Add();
                        IRow rows = sheet.CreateRow(j); //行数
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            if (j == 0) //第一行数据
                            {
                                rows.CreateCell(k).SetCellValue(table.Columns[k].ColumnName.ToString());   //标题栏位
                            }
                            else
                            {
                                if (k > 7)
                                {
                                    rows.CreateCell(k).SetCellValue(Double.Parse(table.Rows[j-1][k].ToString()));   //栏位非标题
                                }
                                else
                                {
                                    rows.CreateCell(k).SetCellValue(table.Rows[j-1][k].ToString());
                                }
                                
                                //rows.CreateCell(k).SetCellValue(int.Parse(table.Rows[j][k].ToString()));   //栏位非标题

                                //if (k>7)
                                //{
                                //    rows.CreateCell(k).SetCellValue(table.Rows[j][k].ToString());   //栏位非标题
                                //}
                                //else
                                //{
                                //    rows.CreateCell(k).SetCellValue(int.Parse(table.Rows[j][k].ToString()));   //栏位非标题
                                //}
                            }
                        }
                    }

                }
                FileStream file = new FileStream(TargetFileNamePath, FileMode.Open, FileAccess.Write);
                workbook.Write(file);
                file.Close();
                return ;
                //Console.ReadKey();
            }

            catch (Exception ex1)
            {
                throw new Exception(ex1.Message);
            }
            finally
            {
                //KillExcelProcess(excel);//结束Excel进程                
                //GC.WaitForPendingFinalizers();
                //GC.Collect();
                //if (sqlConn != null)
                //{
                //    sqlConn.Close();
                //}
            }
        }

    }
}
