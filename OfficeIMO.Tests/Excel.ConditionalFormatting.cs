using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeColor = OfficeIMO.Drawing.OfficeColor;
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
                var workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart wsPart = workbookPart.WorksheetParts.First();
                ConditionalFormatting? cf = wsPart.Worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                Assert.NotNull(cf);
                Assert.Equal("A1:A2", cf!.SequenceOfReferences!.InnerText);
                ConditionalFormattingRule rule = cf.Elements<ConditionalFormattingRule>().First();
                Assert.Equal(ConditionalFormatValues.CellIs, rule.Type!.Value);
                Assert.Equal(ConditionalFormattingOperatorValues.GreaterThan, rule.Operator!.Value);
                Assert.Equal("10", rule.Elements<Formula>().First().Text);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelConditionalFormattingInfo info = Assert.Single(document.Sheets[0].GetConditionalFormattingRules("A1:A3"));
                Assert.Equal("CellIs", info.Type);
                Assert.Equal(nameof(ConditionalFormattingOperatorValues.GreaterThan), info.Operator);
                Assert.Equal(new[] { "10" }, info.Formulas);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_RangeFluentConditionalFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalRangeFluent.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);
                sheet.Range("A1:A3").ConditionalFormatting
                    .ColorScale(OfficeColor.Red, OfficeColor.Lime)
                    .ConditionalFormatting
                    .DataBar(OfficeColor.Blue)
                    .ConditionalFormatting
                    .Top(1)
                    .ConditionalFormatting
                    .DuplicateValues("FCE4D6")
                    .ConditionalFormatting
                    .UniqueValues("DBEAFE")
                    .ConditionalFormatting
                    .AboveAverage("C6EFCE")
                    .ConditionalFormatting
                    .BelowAverage("FEE2E2", equalAverage: true)
                    .ConditionalFormatting
                    .ContainsText("2", "FFF2CC");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var rules = wsPart.Worksheet.Elements<ConditionalFormatting>()
                    .SelectMany(cf => cf.Elements<ConditionalFormattingRule>())
                    .ToList();
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.ColorScale);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.DataBar);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.Top10 && rule.Rank?.Value == 1);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.DuplicateValues && rule.FormatId != null);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.UniqueValues && rule.FormatId != null);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.AboveAverage && rule.AboveAverage?.Value == true && rule.FormatId != null);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.AboveAverage && rule.AboveAverage?.Value == false && rule.EqualAverage?.Value == true && rule.FormatId != null);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.ContainsText && rule.Text?.Value == "2" && rule.FormatId != null);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document.Sheets[0];
                var rules = sheet.GetConditionalFormattingRules("A1:A3");
                ExcelConditionalFormattingInfo colorScale = Assert.Single(rules, info => info.Type == "ColorScale");
                ExcelConditionalFormattingInfo dataBar = Assert.Single(rules, info => info.Type == "DataBar");
                ExcelConditionalFormattingInfo top = Assert.Single(rules, info => info.Type == "Top10");
                ExcelConditionalFormattingInfo duplicate = Assert.Single(rules, info => info.Type == "DuplicateValues");
                ExcelConditionalFormattingInfo unique = Assert.Single(rules, info => info.Type == "UniqueValues");
                ExcelConditionalFormattingInfo above = Assert.Single(rules, info => info.Type == "AboveAverage" && info.AboveAverageAbove);
                ExcelConditionalFormattingInfo below = Assert.Single(rules, info => info.Type == "AboveAverage" && !info.AboveAverageAbove);
                ExcelConditionalFormattingInfo containsText = Assert.Single(rules, info => info.Type == "ContainsText");
                Assert.Equal(new[] { "FFFF0000", "FF00FF00" }, colorScale.ColorScaleColors);
                Assert.Equal("FF0000FF", dataBar.DataBarColor);
                Assert.True(top.Priority > 0);
                Assert.Equal(1U, top.TopBottomRank);
                Assert.False(top.TopBottomBottom);
                Assert.False(top.TopBottomPercent);
                Assert.Equal("FFFCE4D6", duplicate.DifferentialFillColorArgb);
                Assert.Equal("FFDBEAFE", unique.DifferentialFillColorArgb);
                Assert.False(above.AboveAverageEqual);
                Assert.Null(above.AboveAverageStdDev);
                Assert.Equal("FFC6EFCE", above.DifferentialFillColorArgb);
                Assert.True(below.AboveAverageEqual);
                Assert.Null(below.AboveAverageStdDev);
                Assert.Equal("FFFEE2E2", below.DifferentialFillColorArgb);
                Assert.Equal("2", containsText.Text);
                Assert.Equal("ContainsText", containsText.Operator);
                Assert.Equal("FFFFF2CC", containsText.DifferentialFillColorArgb);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AdditionalConditionalFormattingRuleTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalAdditionalRuleTypes.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Ready");
                sheet.CellValue(2, 1, "Blocked");
                sheet.CellValue(3, 1, "Ready");
                sheet.CellValue(1, 2, 1d);
                sheet.CellValue(2, 2, 2d);
                sheet.CellValue(3, 2, 3d);
                sheet.CellValue(1, 3, new System.DateTime(2026, 6, 22));
                sheet.CellValue(2, 3, new System.DateTime(2026, 6, 21));

                sheet.AddConditionalUniqueValuesRule("A1:A3");
                sheet.AddConditionalTextRule("A1:A3", ConditionalFormatValues.ContainsText, "Ready");
                sheet.AddConditionalBlanksRule("D1:D3");
                sheet.AddConditionalErrorsRule("E1:E3", containsErrors: false);
                sheet.AddConditionalAboveAverageRule("B1:B3", aboveAverage: false, equalAverage: true, standardDeviation: 1);
                sheet.AddConditionalTimePeriodRule("C1:C3", TimePeriodValues.Today);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var rules = wsPart.Worksheet.Elements<ConditionalFormatting>()
                    .SelectMany(cf => cf.Elements<ConditionalFormattingRule>())
                    .ToList();
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.UniqueValues);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.ContainsText && rule.Text?.Value == "Ready");
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.ContainsBlanks);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.NotContainsErrors);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.AboveAverage && rule.AboveAverage?.Value == false && rule.EqualAverage?.Value == true && rule.StdDev?.Value == 1);
                Assert.Contains(rules, rule => rule.Type?.Value == ConditionalFormatValues.TimePeriod && rule.TimePeriod?.Value == TimePeriodValues.Today);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var rules = document.Sheets[0].GetConditionalFormattingRules();
                Assert.Contains(rules, info => info.Type == "UniqueValues");
                Assert.Contains(rules, info => info.Type == "ContainsText" && info.Formulas.Count == 1);
                Assert.Contains(rules, info => info.Type == "ContainsBlanks");
                Assert.Contains(rules, info => info.Type == "NotContainsErrors");
                Assert.Contains(rules, info => info.Type == "AboveAverage");
                Assert.Contains(rules, info => info.Type == "TimePeriod" && info.TimePeriod == nameof(TimePeriodValues.Today));
                Assert.Empty(document.ValidateOpenXml());
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
                sheet.AddConditionalColorScale("A1:A3", OfficeColor.Red, OfficeColor.Lime);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart wsPart = workbookPart.WorksheetParts.First();
                ConditionalFormatting? cf = wsPart.Worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                Assert.NotNull(cf);
                ConditionalFormattingRule rule = cf!.Elements<ConditionalFormattingRule>().First();
                Assert.Equal(ConditionalFormatValues.ColorScale, rule.Type!.Value);
                ColorScale? colorScale = rule.GetFirstChild<ColorScale>();
                Assert.NotNull(colorScale);
                var colors = colorScale!.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
                Assert.Equal("FFFF0000", colors[0].Rgb!.Value);
                Assert.Equal("FF00FF00", colors[1].Rgb!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
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
                sheet.AddConditionalDataBar("A1:A3", OfficeColor.Blue);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart wsPart = workbookPart.WorksheetParts.First();
                ConditionalFormatting? cf = wsPart.Worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                Assert.NotNull(cf);
                ConditionalFormattingRule rule = cf!.Elements<ConditionalFormattingRule>().First();
                Assert.Equal(ConditionalFormatValues.DataBar, rule.Type!.Value);
                DataBar? dataBar = rule.GetFirstChild<DataBar>();
                Assert.NotNull(dataBar);
                var color = dataBar!.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().First();
                Assert.Equal("FF0000FF", color.Rgb!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddConditionalIconSet() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalIconSet.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);
                sheet.AddConditionalIconSet("A1:A3", IconSetValues.ThreeTrafficLights1, showValue: false, reverseIconOrder: true);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart wsPart = workbookPart.WorksheetParts.First();
                ConditionalFormatting? cf = wsPart.Worksheet.Elements<ConditionalFormatting>().FirstOrDefault();
                Assert.NotNull(cf);
                ConditionalFormattingRule rule = cf!.Elements<ConditionalFormattingRule>().First();
                Assert.Equal(ConditionalFormatValues.IconSet, rule.Type!.Value);
                IconSet? iconSet = rule.GetFirstChild<IconSet>();
                Assert.NotNull(iconSet);
                Assert.Equal(IconSetValues.ThreeTrafficLights1, iconSet!.IconSetValue!.Value);
                Assert.False(iconSet.ShowValue!.Value);
                Assert.True(iconSet.Reverse!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelConditionalFormattingInfo info = Assert.Single(document.Sheets[0].GetConditionalFormattingRules("A1:A3"));
                Assert.Equal("IconSet", info.Type);
                Assert.Equal("ThreeTrafficLights1", info.IconSet);
                Assert.False(info.IconSetShowValue);
                Assert.True(info.IconSetReverse);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public async Task Test_ConditionalFormattingConcurrent() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalConcurrent.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellValue(3, 1, 3d);

                var tasks = new Task[] {
                    Task.Run(() => sheet.AddConditionalRule("A1:A3", ConditionalFormattingOperatorValues.GreaterThan, "2")),
                    Task.Run(() => sheet.AddConditionalColorScale("A1:A3", OfficeColor.Red, OfficeColor.Blue)),
                    Task.Run(() => sheet.AddConditionalDataBar("A1:A3", OfficeColor.Green))
                };
                await Task.WhenAll(tasks);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                WorksheetPart wsPart = workbookPart.WorksheetParts.First();
                var formats = wsPart.Worksheet.Elements<ConditionalFormatting>().ToList();
                Assert.Contains(formats, cf => cf.Elements<ConditionalFormattingRule>().Any(r => r.Type?.Value == ConditionalFormatValues.CellIs));
                Assert.Contains(formats, cf => cf.Elements<ConditionalFormattingRule>().Any(r => r.Type?.Value == ConditionalFormatValues.ColorScale));
                Assert.Contains(formats, cf => cf.Elements<ConditionalFormattingRule>().Any(r => r.Type?.Value == ConditionalFormatValues.DataBar));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
