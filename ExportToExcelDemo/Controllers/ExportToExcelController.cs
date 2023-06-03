using ExportToExcelDemo.Models;
using ExportToExcelDemo.Models.Services;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace ExportToExcelDemo.Controllers
{
    public class ExportToExcelController : Controller
    {
        List<ExportToExcelModel> dataset = new List<ExportToExcelModel>();

        public ExportToExcelController()
        {
            for (int i = 0; i < 10; i++)
            {
                ExportToExcelModel eTEM = new ExportToExcelModel();

                eTEM.id = i;
                eTEM.column1 = i + 100;
                eTEM.column2 = "Category: " + i;
                eTEM.column3 = i % 2 == 0 ? true : false;
                eTEM.column4 = DateTime.Now.AddDays(i).AddHours(i);
                eTEM.column5 = (short)i;

                dataset.Add(eTEM);
            }
        }

        public IActionResult ExportToExcel()
        {
            return View();
        }

        public ActionResult ExportData()
        {
            if (dataset != null)
            {
                var listByteArray = XlsxConverter.ConvertToXlsx(dataset);
                MemoryStream stream = new MemoryStream(listByteArray);

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "excel_export_" + DateTime.Now.Ticks + ".xlsx");
            }
            
            return View(ExportToExcel());
        }
    }
}
