using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EmployeeExcelSheet.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet("download")]
        public async Task<IActionResult> DownloadExcel()
        {
            // Create a new workbook and worksheet
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                var count = 10;
                worksheet.Range("A1:A2").Merge().Value = "Title 1";
                worksheet.Range("B1:B2").Merge().Value = "Title 2";
                worksheet.Range("C1:C2").Merge().Value = "Title 3" + Environment.NewLine + "A";
                worksheet.Range("D1:D2").Merge().Value = "Title 4" + Environment.NewLine + "B = (K + G + H)";
                worksheet.Range("E1:E2").Merge().Value = "Title 5" + Environment.NewLine + "C = (A - B)";
                worksheet.Range("F1:F2").Merge().Value = "Title 6" + Environment.NewLine + "C = (A - B)";
               

                for (int i = 1; i < count; i++)
                {
                    var rowspan = 2 * i + 1;
                    var nextRowSpan = rowspan + 1;
                    worksheet.Range($"A{rowspan}:A{nextRowSpan}").Merge().Value = 18;
                    worksheet.Range($"B{rowspan}:B{nextRowSpan}").Merge().Value = 2000;
                    worksheet.Range($"C{rowspan}:C{nextRowSpan}").Merge().Value = 2000 + Environment.NewLine + 318.54;
                    worksheet.Cell($"D{rowspan}").Value = 2;
                    worksheet.Cell($"E{rowspan}").Value = 56;
                    worksheet.Cell($"F{rowspan}").Value = 56;
                    worksheet.Cell($"F{rowspan}").Value = 58;

                    worksheet.Cell($"D{nextRowSpan}").Value = 4;
                    worksheet.Cell($"E{nextRowSpan}").Value = 787;
                    worksheet.Cell($"F{nextRowSpan}").Value = 76;
                    worksheet.Cell($"F{nextRowSpan}").Value = 90;
                }
               

                // Adjust column widths and apply formatting
                worksheet.Columns().AdjustToContents();
                worksheet.Range($"A1:F{count * 2}").Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"A1:F{count * 2}").Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                worksheet.Range($"A1:F{count * 2}").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range($"A1:F{count * 2}").Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;


                // Save the workbook to a memory stream
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0;

                    // Return the Excel file as a download
                    return File(
                        stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "practical.xlsx"
                    );
                }
            }
        }
    }
}
