using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTest
{
    public class ExcelReport
    {
        public static async Task<byte[]> GenerateExcel(List<ReportData> reportData)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("info");


            worksheet.Cell(1, 1).Value = "Excel Report";
            var titleStyle = worksheet.Range(1, 1, 1, 5).Merge();
            titleStyle.Style.Font.Bold = true;
            titleStyle.Style.Font.FontSize = 16;
            titleStyle.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var row = 3;
            var col = 1;
            
            // Column headers 
            worksheet.Cell(row, col).Value = "FirstName";
            worksheet.Cell(row, col).Style.Font.Bold = true;
            col += 1;
            worksheet.Cell(row, col).Value = "LastName";
            worksheet.Cell(row, col).Style.Font.Bold = true;
            col += 1;
            worksheet.Cell(row, col).Value = "Date of Birth";
            worksheet.Cell(row, col).Style.Font.Bold = true;
            col += 1;
            worksheet.Cell(row, col).Value = "Age";
            worksheet.Cell(row, col).Style.Font.Bold = true;
            col += 1;
            worksheet.Cell(row, col).Value = "Lucky Number";
            worksheet.Cell(row, col).Style.Font.Bold = true;
            col = 1;


            row = 4;
            foreach (var data in reportData)
            {
                worksheet.Cell(row, col).Value = data.FirstName;
                col += 1;
                worksheet.Cell(row, col).Value = data.LastName;
                col += 1;
                worksheet.Cell(row, col).Value = data.DoB;
                col += 1;
                worksheet.Cell(row, col).Value = data.Age;
                col += 1;
                worksheet.Cell(row, col).Value = data.LuckyNumber;
                col = 1;
                row += 1;
            }


            worksheet.Columns().AdjustToContents();

            MemoryStream memoryStream = new MemoryStream();
            workbook.SaveAs(memoryStream);
            memoryStream.Seek(0, SeekOrigin.Begin);

            var content = memoryStream.ToArray();
            memoryStream.Close();
            return content;
        }


        public class ReportData
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public DateTime DoB { get; set; }
            public int Age { get; set; }
            public double LuckyNumber { get; set; }
        }
    }
}