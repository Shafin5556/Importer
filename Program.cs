using System;
using System.Diagnostics.Metrics;
using System.Drawing;
using System.Runtime.CompilerServices;
using OfficeOpenXml;
using Spire.Pdf;
using Spire.Pdf.Exporting;
namespace routine
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            string path = @"C:\Users\Shafin Ahmed\Desktop\v2.xlsx";
            string searchTerm = "60_PC-B";
            int row, col, pos, counter = 0;
            using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(path)))
            {
                ExcelWorkbook workbook = package.Workbook;
                if (workbook != null)
                {
                    foreach (ExcelWorksheet worksheet in workbook.Worksheets)
                    {
                        for (row = 1; row <= worksheet.Dimension.End.Row; row++)
                        {
                            for (col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {

                                string cellValue = worksheet.Cells[row, col].Value?.ToString() ?? string.Empty;
                                if (col > 1)
                                {

                                    pos = col - 1;
                                    string room = worksheet.Cells[row, pos].Value?.ToString() ?? string.Empty;
                                    string teacher = worksheet.Cells[row, (col + 1)].Value?.ToString() ?? string.Empty;
                                    if (cellValue.Contains(searchTerm))
                                    {
                                        Console.WriteLine($"{cellValue} with {teacher} in {room}");
                                    }

                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
