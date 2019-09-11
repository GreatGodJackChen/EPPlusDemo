using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace EPPlusDemo.Excel
{
    public class ExcelHelper
    {
        public static MemoryStream ExportExcel(DataTable dt)
        {
           return ExportExcel(dt, new string[10000000]);
        }

        public static MemoryStream ExportExcel(DataTable dt, string[] columnName)
        {
            try
            {
                if (dt == null || dt.Rows.Count <= 0)
                {
                    return null;
                }
                //定义数据流
                var stream = new MemoryStream();
                //导出文件
                using (ExcelPackage package = new ExcelPackage(stream))
                {
                    //添加worksheet
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("测试Excel");
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
                stream.Position = 0;
                return stream;
            }
            catch (Exception ex)
            {
                //写入日志
                return null;
            }
        }
    }
}
