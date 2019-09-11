using Microsoft.AspNetCore.Hosting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusDemo.Common
{
    public class ExcelHelper
    {
        ////private readonly IHostingEnvironment _hostingEnvironment;
        //public ExportExcelHelper(IHostingEnvironment hostingEnvironment)
        //{
        //    //_hostingEnvironment = hostingEnvironment;
        //}
        #region Excel导出
        /// <summary>
        /// 导出数据Excel
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="msg">消息</param>
        /// <returns></returns>
        public static bool ExportExcel(DataTable dt, ref string msg)
        {
            return ExportExcel(dt, ref msg, new string[1000000], "", "", 0, 0);
        }
        /// <summary>
        /// 导出数据Excel
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="msg">消息</param>
        /// <param name="columnName">列名</param>
        /// <returns></returns>
        public static bool ExportExcel(DataTable dt, ref string msg, string[] columnName)
        {
            return ExportExcel(dt, ref msg, columnName, "", "", 0, 0);
        }
        /// <summary>
        /// 导出数据Excel
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="msg">消息</param>
        /// <param name="columnName">列名</param>
        /// <param name="fileName">文件名</param>
        /// <param name="rootFolder">文件夹</param>
        /// <returns></returns>
        public static bool ExportExcel(DataTable dt, ref string msg, string[] columnName, string fileName, string rootFolder)
        {
            return ExportExcel(dt, ref msg, columnName, fileName, rootFolder, 0, 0);
        }
        /// <summary>
        /// 导出数据Excel
        /// </summary>
        /// <param name="dt">数据源</param>
        /// <param name="msg">消息</param>
        /// <param name="columnName">列名</param>
        /// <param name="fileName">文件名</param>
        /// <param name="rootFolder">文件夹</param>
        /// <param name="fileType">文件名类型：0文件名+时间，其他为时间类型</param>
        /// <param name="rootType">文件夹类型</param>
        /// <returns></returns>
        public static bool ExportExcel(DataTable dt, ref string msg, string[] columnName, string fileName = "", string rootFolder = "", int fileType = 0, int rootType = 0)
        {
            try
            {
                if (dt == null || dt.Rows.Count <= 0)
                {
                    msg = "没有找到数据";
                    return false;
                }
                //文件类型
                switch (fileType)
                {
                    case 0:
                        //按文件名时间生成excel文件
                        fileName = fileName + DateTime.Now.ToString("yyyymmddhhmmssfff");
                        break;
                    default:
                        fileName = DateTime.Now.ToString("yyyymmddhhmmssfff");
                        break;
                }
                //utf-8转换
                UTF8Encoding utf8 = new UTF8Encoding();
                byte[] buffer = utf8.GetBytes(fileName);
                fileName = Encoding.UTF8.GetString(buffer);
                //文件夹类型
                switch (rootType)
                {
                    case 0:
                        rootFolder = Directory.GetCurrentDirectory() + "/" + rootFolder + DateTime.Now.ToString("yyyymmdd");
                        break;
                    default:
                        rootFolder = Directory.GetCurrentDirectory() + "/" + DateTime.Now.ToString("yyyymmdd");
                        break;
                }
                //判断文件夹              
                if (!Directory.Exists(rootFolder))
                {
                    Directory.CreateDirectory(rootFolder);
                }
                //判断是否有相同文件名
                FileInfo file = new FileInfo(Path.Combine(rootFolder, fileName));
                if (file.Exists)
                {
                    //判断同名文件创建时间
                    file.Delete();
                    file = new FileInfo(Path.Combine(rootFolder, fileName));
                }
                //导出文件
                using (ExcelPackage package = new ExcelPackage(file))
                {
                    //添加worksheet
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(fileName.Split('.')[0]);
                    //添加表头
                    int column = 1;
                    if (columnName.Count() == dt.Columns.Count)
                    {
                        foreach (string cn in columnName)
                        {
                            worksheet.Cells[1, column].Value = cn.Trim();

                            worksheet.Cells[1, column].Style.Font.Bold = true;//字体为粗体
                            worksheet.Cells[1, column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;//水平居中
                            worksheet.Cells[1, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;//设置样式类型
                            worksheet.Cells[1, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(159, 197, 232));//设置单元格背景色
                            column++;
                        }
                    }
                    else
                    {
                        foreach (DataColumn dc in dt.Columns)
                        {
                            worksheet.Cells[1, column].Value = dc.ColumnName;

                            worksheet.Cells[1, column].Style.Font.Bold = true;//字体为粗体
                            worksheet.Cells[1, column].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;//水平居中
                            worksheet.Cells[1, column].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;//设置样式类型
                            worksheet.Cells[1, column].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(159, 197, 232));//设置单元格背景色
                            column++;
                        }
                    }
                    //添加数据
                    int row = 2;
                    foreach (DataRow dr in dt.Rows)
                    {
                        int col = 1;
                        foreach (DataColumn dc in dt.Columns)
                        {
                            worksheet.Cells[row, col].Value = dr[col - 1].ToString();
                            col++;
                        }
                        row++;
                    }
                    //自动列宽
                    worksheet.Cells.AutoFitColumns();

                    //保存workbook.
                    package.Save();
                }
                return true;
            }
            catch (Exception ex)
            {
                msg = "生成Excel失败：" + ex.Message;
                //写入日志
                return false;
            }
        }
        #endregion

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="filepath">文件路径</param>
        /// <param name="msg">消息</param>
        /// <returns></returns>
        public static List<T> ImportExcel<T>(string filepath,ref string msg) where T : new()
        {
            try
            {
                //定义返回集合
                List<T> result=new List<T>();
                FileInfo file = new FileInfo(filepath);
                if (file != null)
                {
                    using (ExcelPackage package = new ExcelPackage(file))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
                        result = worksheet.ConvertSheetToObjects<T>().ToList();
                    }
                }
                return result;
            }
            catch(Exception ex)
            {
                msg = "导入Excel失败：" + ex.Message;
                //写入日志
                return null;
            }
        }
        public static DataTable GetDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Id", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("描述", typeof(string));
            dt.Columns.Add("性别", typeof(string));
            dt.Columns.Add("年龄", typeof(Int32));
            for (int i = 0; i < 40000; i++)
            {
                dt.Rows.Add(new Object[] {Guid.NewGuid(),"cj","dddddddddddssf发射点发生发生发射点","男",22 });
            }
            return dt;
        }
    }
}
