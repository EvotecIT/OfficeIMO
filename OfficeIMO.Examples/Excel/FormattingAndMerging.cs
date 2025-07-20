using System;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates formatting and merging cells.
    /// </summary>
    public static class FormattingAndMerging {
        public static void Run(string folderPath, bool openExcel) {
            var filePath = Path.Combine(folderPath, "FormattingAndMerging.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sheet1");

                var cell = sheet.GetCell("A1");
                cell.Text = "Merged";
                cell.Font = new Font(new Bold());
                cell.Fill = new Fill(new PatternFill(new ForegroundColor { Rgb = "FFFF00" }) { PatternType = PatternValues.Solid });
                cell.Border = new Border(new LeftBorder(new Color { Rgb = "FF0000" }) { Style = BorderStyleValues.Thin },
                    new RightBorder(new Color { Rgb = "FF0000" }) { Style = BorderStyleValues.Thin },
                    new TopBorder(new Color { Rgb = "FF0000" }) { Style = BorderStyleValues.Thin },
                    new BottomBorder(new Color { Rgb = "FF0000" }) { Style = BorderStyleValues.Thin },
                    new DiagonalBorder());
                cell.NumberFormat = "@";

                sheet.MergeCells("A1:C1");

                document.Save(openExcel);
            }
        }
    }
}
