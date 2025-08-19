using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AddConditionalRule() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalRule.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, 5d);
                sheet.SetCellValue(2, 1, 15d);
                sheet.AddConditionalRule("A1:A2", ConditionalFormattingOperatorValues.GreaterThan, "10");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                ConditionalFormatting cf = wsPart.Worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                Assert.NotNull(cf);
                Assert.Equal("A1:A2", cf.SequenceOfReferences.InnerText);
                ConditionalFormattingRule rule = cf.Elements<ConditionalFormattingRule>().First();
                Assert.Equal(ConditionalFormatValues.CellIs, rule.Type.Value);
                Assert.Equal(ConditionalFormattingOperatorValues.GreaterThan, rule.Operator.Value);
                Assert.Equal("10", rule.Elements<Formula>().First().Text);
            }
        }

        [Fact]
        public void Test_AddConditionalColorScale() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalColorScale.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, 1d);
                sheet.SetCellValue(2, 1, 2d);
                sheet.SetCellValue(3, 1, 3d);
                sheet.AddConditionalColorScale("A1:A3", "FFFF0000", "FF00FF00");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                ConditionalFormatting cf = wsPart.Worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                Assert.NotNull(cf);
                ConditionalFormattingRule rule = cf.Elements<ConditionalFormattingRule>().First();
                Assert.Equal(ConditionalFormatValues.ColorScale, rule.Type.Value);
                ColorScale colorScale = rule.GetFirstChild<ColorScale>();
                Assert.NotNull(colorScale);
                var colors = colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                Assert.Equal("FFFF0000", colors[0].Rgb.Value);
                Assert.Equal("FF00FF00", colors[1].Rgb.Value);
            }
        }
    }
}
