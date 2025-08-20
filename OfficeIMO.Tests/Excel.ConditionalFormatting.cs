using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using SixLaborsColor = SixLabors.ImageSharp.Color;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AddConditionalRule() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalRule.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 5d);
                sheet.CellValue(2, 1, 15d);
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
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);
                sheet.AddConditionalColorScale("A1:A3", SixLaborsColor.Red, SixLaborsColor.Lime);
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

        [Fact]
        public void Test_AddConditionalDataBar() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalDataBar.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);
                sheet.AddConditionalDataBar("A1:A3", SixLaborsColor.Blue);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                ConditionalFormatting cf = wsPart.Worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                Assert.NotNull(cf);
                ConditionalFormattingRule rule = cf.Elements<ConditionalFormattingRule>().First();
                Assert.Equal(ConditionalFormatValues.DataBar, rule.Type.Value);
                DataBar dataBar = rule.GetFirstChild<DataBar>();
                Assert.NotNull(dataBar);
                var color = dataBar.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().First();
                Assert.Equal("FF0000FF", color.Rgb.Value);
            }
        }

        [Fact]
        public void Test_ConditionalFormattingConcurrent() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalConcurrent.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);

                var tasks = new Task[] {
                    Task.Run(() => sheet.AddConditionalRule("A1:A3", ConditionalFormattingOperatorValues.GreaterThan, "2")),
                    Task.Run(() => sheet.AddConditionalColorScale("A1:A3", SixLaborsColor.Red, SixLaborsColor.Blue)),
                    Task.Run(() => sheet.AddConditionalDataBar("A1:A3", SixLaborsColor.Green))
                };
                Task.WaitAll(tasks);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                var formats = wsPart.Worksheet.Elements<ConditionalFormatting>().ToList();
                Assert.Contains(formats, cf => cf.Elements<ConditionalFormattingRule>().Any(r => r.Type == ConditionalFormatValues.CellIs));
                Assert.Contains(formats, cf => cf.Elements<ConditionalFormattingRule>().Any(r => r.Type == ConditionalFormatValues.ColorScale));
                Assert.Contains(formats, cf => cf.Elements<ConditionalFormattingRule>().Any(r => r.Type == ConditionalFormatValues.DataBar));
            }
        }
    }
}
