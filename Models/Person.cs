using EPPlusDemo.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlusDemo.Models
{
    public class Person
    {
        [ExcelColumn("姓名")]
        public string Name { get; set; }
        [ExcelColumn("编号")]
        public int Id { get; set; }
        [ExcelColumn("年龄")]
        public int Age { get; set; }
        [ExcelColumn("性别")]
        public string Sex { get; set; }
        [ExcelColumn("描述")]
        public string describe{ get; set; }
    }
}
