using OfficeOpenXml;
using System.IO;

namespace PineWealth
{
    /// <summary>
    /// Exports translation data to an Excel spreadsheet
    /// </summary>
    public class ExcelExporter
    {
        /// <summary>
        /// Exports the Pine Script source, C# translation, and related data to an Excel file
        /// </summary>
        /// <param name="pineScriptSource">The original Pine Script source code</param>
        /// <param name="csharpCode">The translated C# code</param>
        /// <param name="filePath">The file path to save the Excel file</param>
        public void Export(string pineScriptSource, string csharpCode, string filePath)
        {
            // Set EPPlus license context (non-commercial personal use)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                // Create Code Comparison sheet
                CreateCodeComparisonSheet(package, pineScriptSource, csharpCode);

                // Create Pine Script Lines sheet
                CreatePineScriptSheet(package, pineScriptSource);

                // Create C# Code sheet
                CreateCSharpSheet(package, csharpCode);

                // Save the file
                var fileInfo = new FileInfo(filePath);
                package.SaveAs(fileInfo);
            }
        }

        /// <summary>
        /// Creates a sheet showing Pine Script and C# code side by side
        /// </summary>
        private void CreateCodeComparisonSheet(ExcelPackage package, string pineScriptSource, string csharpCode)
        {
            var worksheet = package.Workbook.Worksheets.Add("Code Comparison");

            // Headers
            worksheet.Cells[1, 1].Value = "Line #";
            worksheet.Cells[1, 2].Value = "Pine Script";
            worksheet.Cells[1, 3].Value = "C# Code";

            // Style headers
            using (var range = worksheet.Cells[1, 1, 1, 3])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }

            // Split source into lines
            var pineLines = pineScriptSource.Replace("\r", "").Split('\n');
            var csharpLines = csharpCode.Replace("\r", "").Split('\n');

            // Add Pine Script lines
            int maxRows = Math.Max(pineLines.Length, csharpLines.Length);
            for (int i = 0; i < maxRows; i++)
            {
                worksheet.Cells[i + 2, 1].Value = i + 1;
                if (i < pineLines.Length)
                    worksheet.Cells[i + 2, 2].Value = pineLines[i];
                if (i < csharpLines.Length)
                    worksheet.Cells[i + 2, 3].Value = csharpLines[i];
            }

            // Auto-fit columns
            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 60;
            worksheet.Column(3).Width = 80;
        }

        /// <summary>
        /// Creates a sheet containing only the Pine Script source
        /// </summary>
        private void CreatePineScriptSheet(ExcelPackage package, string pineScriptSource)
        {
            var worksheet = package.Workbook.Worksheets.Add("Pine Script");

            // Headers
            worksheet.Cells[1, 1].Value = "Line #";
            worksheet.Cells[1, 2].Value = "Code";

            // Style headers
            using (var range = worksheet.Cells[1, 1, 1, 2])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            }

            // Split source into lines
            var lines = pineScriptSource.Replace("\r", "").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                worksheet.Cells[i + 2, 1].Value = i + 1;
                worksheet.Cells[i + 2, 2].Value = lines[i];
            }

            // Auto-fit columns
            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 80;
        }

        /// <summary>
        /// Creates a sheet containing only the C# translated code
        /// </summary>
        private void CreateCSharpSheet(ExcelPackage package, string csharpCode)
        {
            var worksheet = package.Workbook.Worksheets.Add("C# Code");

            // Headers
            worksheet.Cells[1, 1].Value = "Line #";
            worksheet.Cells[1, 2].Value = "Code";

            // Style headers
            using (var range = worksheet.Cells[1, 1, 1, 2])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
            }

            // Split code into lines
            var lines = csharpCode.Replace("\r", "").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                worksheet.Cells[i + 2, 1].Value = i + 1;
                worksheet.Cells[i + 2, 2].Value = lines[i];
            }

            // Auto-fit columns
            worksheet.Column(1).Width = 10;
            worksheet.Column(2).Width = 100;
        }
    }
}
