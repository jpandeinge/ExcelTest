using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using ExcelTest;
using Xunit;
using static System.IO.FileAccess;

namespace Tests
{
    public class Report
    {
        [Fact]
        public async Task TestGenerateExcelReport()
        {
            var data = AddDummyData(10);
            var report = await ExcelReport.GenerateExcel(data);

            string filename = "Excel Report";
            FileStream file = new FileStream($"C:/temp/{filename}.xlsx", FileMode.Create);
            file.Write(report);
            file.Close();
        }

        public List<ExcelReport.ReportData> AddDummyData(int numOfRows)
        {
            var reportData = new List<ExcelReport.ReportData>();

            for (int i = 0; i < numOfRows; i++)
            {
                var temp = new ExcelReport.ReportData
                {
                    FirstName = $"FirstName {i}",
                    LastName = $"LastName {i}",
                    DoB = new DateTime(2021, 07, 1),
                    Age = i,
                    LuckyNumber = i
                };
                reportData.Add(temp);
            }
            return reportData;
        }
    }
}