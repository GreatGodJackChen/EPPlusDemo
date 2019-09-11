using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EPPlusDemo.Common;
using EPPlusDemo.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EPPlusDemo.Controllers
{
    public class ExcelController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult Import(IFormFile excelfile)
        {
            string msg = "";
            var file = excelfile.FileName;
            var result=  ExcelHelper.ImportExcel<Person>(file,ref msg);
            return View();
        }
        public IActionResult Export()
        {
            string msg = "";
            var dt = ExcelHelper.GetDataTable();
            Response.ContentType = "application/vnd.ms-excel";
            var result = ExcelHelper.ExportExcel(dt,ref msg);
            return View();
        }
        public IActionResult Export2()
        {
            //指定默认文件名
            var time = DateTime.Now.ToString("yyyymmddhhmmssfff");
            string fileName = $"{time}.xlsx";
            var dt = ExcelHelper.GetDataTable();
            var result = EPPlusDemo.Excel.ExcelHelper.ExportExcel(dt);
            return File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",fileName);
        }
    }
}