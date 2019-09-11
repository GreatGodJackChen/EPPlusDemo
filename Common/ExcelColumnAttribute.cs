using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlusDemo.Common
{
    /// <summary>
    /// 自定义excel头部标签
    /// </summary>
    [AttributeUsage(AttributeTargets.All)]
    public class ExcelColumnAttribute: Attribute
    {
        // <summary>
        /// 标签名称
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="name"></param>
        public ExcelColumnAttribute(string name)
        {
            ColumnName = name;
        }
    }
}
