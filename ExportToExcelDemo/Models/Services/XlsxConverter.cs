using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace ExportToExcelDemo.Models.Services
{
    public static class XlsxConverter
    {
        public static byte[] ConvertToXlsx<T>(List<T> data)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                worksheet.Cells["A1"].LoadFromCollection(data, true);

                return excelPackage.GetAsByteArray();
            }
        }

        public static byte[] ConvertToXlsx(List<dynamic> dataset)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                DataTable dataTable = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(dataset), (typeof(DataTable)));

                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                int colNumber = 0;
                foreach (DataColumn col in dataTable.Columns)
                {
                    colNumber++;
                    if (col.DataType == typeof(DateTime)) worksheet.Column(colNumber).Style.Numberformat.Format = "yyyy-MM-dd hh:mm:ss";
                }

                return excelPackage.GetAsByteArray();
            }
        }
    }
}
