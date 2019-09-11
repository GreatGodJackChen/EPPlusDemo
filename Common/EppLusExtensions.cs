using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EPPlusDemo.Common
{
    public static class EppLusExtensions
    {
        /// <summary>
        /// 获取标签对应excel的Index
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static int GetColumnByName(this ExcelWorksheet ws, string columnName)
        {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            return ws.Cells["1:1"].First(c => c.Value.ToString() == columnName).Start.Column;
        }
        /// <summary>
        /// 扩展方法
        /// </summary>
        /// <param name="worksheet"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet) where T : new()
        {

            Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(ExcelColumnAttribute);
            var columns = typeof(T)
                .GetProperties()
                .Where(x => x.CustomAttributes.Any(columnOnly))
                .Select(p => new
                {
                    Property = p,
                    p.GetCustomAttributes<ExcelColumnAttribute>().First().ColumnName
                }).ToList();

            var rows = worksheet.Cells
                .Select(cell => cell.Start.Row)
                .Distinct()
                .OrderBy(x => x);

            var collection = rows.Skip(1)
                .Select(row =>
                {
                    var tnew = new T();
                    columns.ForEach(col =>
                    {
                        var val = worksheet.Cells[row, GetColumnByName(worksheet, col.ColumnName)];
                        if (val.Value == null)
                        {
                            col.Property.SetValue(tnew, null);
                            return;
                        }
                        // 如果类的对应字段是int的，该怎么怎么做……
                        if (col.Property.PropertyType == typeof(int))
                        {
                            col.Property.SetValue(tnew, val.GetValue<int>());
                            return;
                        }
                        // 如果类的对应字段是double的，该怎么怎么做……
                        if (col.Property.PropertyType == typeof(double))
                        {
                            col.Property.SetValue(tnew, val.GetValue<double>());
                            return;
                        }
                        // 如果类的对应字段是DateTime?的，该怎么怎么做……
                        if (col.Property.PropertyType == typeof(DateTime?))
                        {
                            col.Property.SetValue(tnew, val.GetValue<DateTime?>());
                            return;
                        }
                        // 如果类的对应字段是DateTime的，该怎么怎么做……
                        if (col.Property.PropertyType == typeof(DateTime))
                        {
                            col.Property.SetValue(tnew, val.GetValue<DateTime>());
                            return;
                        }
                        // 如果类的对应字段是bool的，该怎么怎么做……
                        if (col.Property.PropertyType == typeof(bool))
                        {
                            col.Property.SetValue(tnew, val.GetValue<bool>());
                            return;
                        }
                        col.Property.SetValue(tnew, val.GetValue<string>());
                    });

                    return tnew;
                });
            return collection;
        }
    }
}
