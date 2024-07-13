using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        string jsonFilePath = @"C:\Users\gnoor\source\repos\tassk\tassk\783_46622633507162.json";
        string excelFilePath = @"C:\Users\gnoor\source\repos\tassk\tassk\output.xlsx";
        string jsonContent = File.ReadAllText(jsonFilePath);
        var jsonArray = JsonConvert.DeserializeObject<JArray>(jsonContent);

        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets.Add("Data");

            // نوشتن سرستون‌های اصلی با نام‌های فارسی
            worksheet.Cells[1, 1].Value = "نام فرآیند";
            worksheet.Cells[1, 2].Value = "واحد سنجش";
            worksheet.Cells[1, 3].Value = "مرکز شماره 1 اراک";

            // استایل‌دهی به سرستون‌ها
            using (var range = worksheet.Cells[1, 1, 1, 3])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            int row = 2;
            foreach (var item in jsonArray)
            {
                string processName = item["process_name"].ToString();
                var company = item["companies"][0];
                string companyName = company["company_name"].ToString();
                string totalFrequency = company["total_frequency"].ToString();
                string averageExecutionTime = company["average_execution_time"].ToString();
                string totalCalculations = company["total_calculations"].ToString();

                worksheet.Cells[row, 1].Value = processName;
                worksheet.Cells[row, 2].Value = "total_frequency";
                worksheet.Cells[row, 3].Value = totalFrequency;
                worksheet.Cells[row + 1, 2].Value = "average_execution_time";
                worksheet.Cells[row + 1, 3].Value = averageExecutionTime;
                worksheet.Cells[row + 2, 2].Value = "total_calculations";
                worksheet.Cells[row + 2, 3].Value = totalCalculations;

                // ادغام سلول‌های process_name
                worksheet.Cells[row, 1, row + 2, 1].Merge = true;
                worksheet.Cells[row, 1, row + 2, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[row, 1, row + 2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                row += 3;
            }

            // تنظیم عرض ستون "نام فرآیند" به صورت دستی
            worksheet.Column(1).Width = 20; // تنظیم عرض به 40 واحد

            // تنظیم عرض سایر ستون‌ها به صورت خودکار
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).AutoFit();

            // اضافه کردن خطوط به جدول
            var tableRange = worksheet.Cells[1, 1, row - 1, 3];
            tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

            // تنظیم تراز عمودی و افقی برای همه سلول‌ها
            tableRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            tableRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            // فعال کردن قابلیت wrap text برای ستون "نام فرآیند"
            worksheet.Column(1).Style.WrapText = true;

            // ذخیره فایل Excel
            package.Save();
        }

        Console.WriteLine("Excel file has been created successfully.");
    }
}