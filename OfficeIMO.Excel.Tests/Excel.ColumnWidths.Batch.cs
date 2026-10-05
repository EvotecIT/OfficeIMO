using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(12D)]
        [InlineData(0D)]
        public void Test_SetColumnWidth_SplitsRangeWithoutChangingNeighborWidthsOrFitMetadata(double width) {
            string filePath = Path.Combine(_directoryWithFiles, "ColumnWidths.SingleRange." + width + ".xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 2, "A label that must grow beyond twelve characters when fitted");
                var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet;
                worksheet.InsertAt(new Columns(new Column {
                    Min = 1, Max = 3, Width = 30D, CustomWidth = true, BestFit = true,
                    Hidden = true, OutlineLevel = 2
                }), 0);
                sheet.SetColumnWidth(2, width);
                var columns = worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
                Assert.Equal(new uint[] { 1, 2, 3 }, columns.Select(column => column.Min!.Value));
                Assert.Equal(new double?[] { 30D, width > 0D ? width : null, 30D }, columns.Select(column => column.Width?.Value));
                Assert.True(columns[0].BestFit!.Value && columns[2].BestFit!.Value);
                Assert.NotEqual(true, columns[1].BestFit?.Value);
                Assert.All(columns, column => {
                    Assert.True(column.Hidden!.Value);
                    Assert.Equal((byte)2, column.OutlineLevel!.Value);
                });
                document.Save();
            }
            using var reopened = ExcelDocument.Load(filePath);
            reopened.Sheets[0].AutoFitColumnsFor(new[] { 2 });
            var saved = reopened.Sheets[0].GetColumnDefinitions().ToArray();
            Assert.Equal(30D, saved[0].Width);
            Assert.True(saved[1].Width > 12D);
            Assert.Equal(30D, saved[2].Width);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void Test_ManualColumnWidths_DoNotPreventAutoFitAfterReopening(bool batch, bool previouslyFitted) {
            string filePath = Path.Combine(_directoryWithFiles, "ColumnWidths.Refit." + batch + "." + previouslyFitted + ".xlsx");
            const string label = "A long label whose content needs more than twelve characters of column width";
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, label);
                if (previouslyFitted) sheet.AutoFitColumns();
                if (batch) sheet.SetColumnWidths(new Dictionary<int, double> { [1] = 12D });
                else sheet.SetColumnWidth(1, 12D);
                document.Save();
            }
            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets[0];
                Assert.Equal(12D, Assert.Single(sheet.GetColumnDefinitions()).Width);
                sheet.AutoFitColumnsFor(new[] { 1 });
                Assert.True(Assert.Single(sheet.GetColumnDefinitions()).Width > 12D);
                document.Save();
            }
            using var reopened = ExcelDocument.Load(filePath);
            Assert.Equal(label, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.True(Assert.Single(reopened.Sheets[0].GetColumnDefinitions()).Width > 12D);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_SetColumnWidths_SavesOnceAndPreservesColumnMetadata(bool withRange) {
            string filePath = Path.Combine(_directoryWithFiles, "ColumnWidths.Batch." + withRange + ".xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Data");
                for (int column = 1; column <= 64; column++) sheet.CellValue(1, column, "Cell " + column);
                sheet.AutoFitColumns();
                var package = document.OpenXmlDocument;
                var worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                var worksheet = worksheetPart.Worksheet;
                // Include an authored range alongside the singleton definitions created by auto-fit.
                if (withRange) worksheet.GetFirstChild<Columns>()!.Append(new Column {
                    Min = 65, Max = 68, Width = 10D, CustomWidth = true,
                    Hidden = true, OutlineLevel = 2, Collapsed = true
                });
                package.AddPartRootEventsFeature();
                int saves = 0;
                package.Features.Get<IPartRootEventsFeature>()!.Change += args => {
                    if (args.Type == EventType.Saved && ReferenceEquals(args.Argument, worksheetPart)) saves++;
                };
                var widths = Enumerable.Range(1, 64).ToDictionary(column => column, column => 12D + column / 10D);
                if (withRange) widths.Add(66, 25D);
                sheet.SetColumnWidths(widths);
                Assert.Equal(1, saves);
                if (withRange) {
                    ExcelColumnSnapshot selected = sheet.GetColumnDefinitions().Single(column => column.StartIndex == 66);
                    Assert.Equal(25D, selected.Width);
                    Assert.True(selected.Hidden && selected.Collapsed);
                    Assert.Equal((byte)2, selected.OutlineLevel);
                    Assert.Equal(10D, sheet.GetColumnDefinitions().Single(column => column.StartIndex == 65).Width);
                    Assert.Equal(10D, sheet.GetColumnDefinitions().Single(column => column.StartIndex == 67).Width);
                }
                document.Save();
            }
            using var reopened = ExcelDocument.Load(filePath);
            Assert.Equal("Cell 64", reopened.Sheets[0].CellAt(1, 64).GetValue<string>());
            Assert.Equal(18.4D, reopened.Sheets[0].GetColumnDefinitions().Single(column => column.StartIndex == 64).Width);
        }

        [Theory]
        [InlineData(0, 12D)]
        [InlineData(16385, 12D)]
        [InlineData(2, 0D)]
        [InlineData(2, double.NaN)]
        [InlineData(2, double.PositiveInfinity)]
        public void Test_SetColumnWidths_ValidatesTheWholeBatchBeforeChangingWidths(int invalidIndex, double invalidWidth) {
            using var document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "ColumnWidths.Invalid.xlsx"));
            var sheet = document.AddWorksheet("Data");
            sheet.SetColumnWidth(1, 20D);
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetColumnWidths(new Dictionary<int, double> {
                [1] = 30D, [invalidIndex] = invalidWidth
            }));
            Assert.Equal(20D, Assert.Single(sheet.GetColumnDefinitions()).Width);
        }
    }
}
