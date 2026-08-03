using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ModernChart_UpdateDataAllocatesAwayFromSharedChartPartRange() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A", "B" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }) }),
                1,
                1,
                ExcelModernChartType.Funnel);
            ExcelChartDataRange sharedRange = chart.DataRange!;
            DrawingsPart drawingsPart = sheet.WorksheetPart.DrawingsPart!;
            ExtendedChartPart originalPart = Assert.Single(drawingsPart.ExtendedChartParts);
            ExtendedChartPart sharedPart = drawingsPart.AddNewPart<ExtendedChartPart>();
            sharedPart.ChartSpace = (Cx.ChartSpace)originalPart.ChartSpace!.CloneNode(true);
            sharedPart.ChartSpace.Save();

            Xdr.OneCellAnchor duplicate = (Xdr.OneCellAnchor)drawingsPart.WorksheetDrawing!
                .Elements<Xdr.OneCellAnchor>().Single().CloneNode(true);
            duplicate.Descendants<Xdr.NonVisualDrawingProperties>().Single().Id = 99U;
            duplicate.Descendants<Xdr.NonVisualDrawingProperties>().Single().Name = "Shared data chart";
            duplicate.Descendants<Cx.RelId>().Single().Id = drawingsPart.GetIdOfPart(sharedPart);
            drawingsPart.WorksheetDrawing.Append(duplicate);
            drawingsPart.WorksheetDrawing.Save();

            chart.UpdateData(new ExcelChartData(
                new[] { "A", "B" },
                new[] { new ExcelChartSeries("Updated", new[] { 9d, 10d }) }));

            Assert.NotEqual(sharedRange.DataRangeA1, chart.DataRange!.DataRangeA1);
            ExcelSheet dataSheet = document[sharedRange.SheetName];
            Assert.Equal(1d, dataSheet.CellAt(sharedRange.SeriesStartRow, sharedRange.SeriesStartColumn).GetValue<double>());
            Assert.Equal(2d, dataSheet.CellAt(sharedRange.SeriesStartRow + 1, sharedRange.SeriesStartColumn).GetValue<double>());
            Assert.Contains(sheet.ModernCharts, item =>
                item.DataRange?.DataRangeA1 == sharedRange.DataRangeA1);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_ModernChart_MalformedDrawingRelationshipFailsBeforeBackingDataMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            sheet.WorksheetPart.Worksheet.Append(
                new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = "rIdMissing" });
            sheet.WorksheetPart.Worksheet.Save();

            Assert.ThrowsAny<Exception>(() => sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel));

            Assert.Single(document.Sheets);
            Assert.Null(sheet.WorksheetPart.DrawingsPart);
        }

        [Theory]
        [InlineData("1:1", "A1:XFD1", ExcelCellShiftDirection.Down, true, "2:2")]
        [InlineData("2:2", "A1:XFD1", ExcelCellShiftDirection.Up, false, "1:1")]
        [InlineData("1:1", "A1:XFD1", ExcelCellShiftDirection.Up, false, null)]
        [InlineData("A:A", "A1:A1048576", ExcelCellShiftDirection.Right, true, "B:B")]
        [InlineData("B:B", "A1:A1048576", ExcelCellShiftDirection.Left, false, "A:A")]
        [InlineData("A:A", "A1:A1048576", ExcelCellShiftDirection.Left, false, null)]
        public void Test_CellShiftReference_FullBandsRemapWholeRowsAndColumns(
            string reference,
            string affected,
            ExcelCellShiftDirection direction,
            bool inserting,
            string? expected) {
            ExcelReference? transformed = ExcelDocument.TransformCellShiftReference(
                ExcelReference.Parse(reference),
                ExcelReference.Parse(affected),
                direction,
                inserting);

            Assert.Equal(expected, transformed?.ToString());
        }

        [Fact]
        public void Test_RangeMove_ReplacesDestinationRangeMetadataBeforeMovingSourceRules() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(1, 3, 9);
            sheet.ValidationList("A1", new[] { "source" });
            sheet.ValidationList("C1", new[] { "destination" });
            sheet.AddConditionalFormulaRule("A1", "A1=1");
            sheet.AddConditionalFormulaRule("C1", "C1=9");
            sheet.Protect();
            sheet.SetAllowedEditRange("Source", new[] { "A1" });
            sheet.SetAllowedEditRange("Destination", new[] { "C1" });
            sheet.AddIgnoredErrorRegion(new[] { "A1" }, ExcelIgnoredErrorKind.NumberStoredAsText);
            sheet.AddIgnoredErrorRegion(new[] { "C1" }, ExcelIgnoredErrorKind.Formula);

            sheet.MoveRange("A1", "C1");

            ExcelDataValidationInfo validation = Assert.Single(sheet.GetDataValidations());
            Assert.Equal("C1", validation.Range);
            Assert.Contains("source", validation.Formula1, StringComparison.Ordinal);
            ExcelConditionalFormattingInfo formatting = Assert.Single(sheet.GetConditionalFormattingRules());
            Assert.Equal("C1", formatting.Range);
            Assert.Contains("C1=1", formatting.Formulas);
            ExcelAllowedEditRangeInfo allowed = Assert.Single(sheet.GetAllowedEditRanges());
            Assert.Equal("Source", allowed.Name);
            Assert.Equal(new[] { "C1" }, allowed.Ranges);
            ExcelIgnoredErrorRegionInfo ignored = Assert.Single(sheet.GetIgnoredErrorRegions());
            Assert.Equal(new[] { "C1" }, ignored.Ranges);
            Assert.Equal(ExcelIgnoredErrorKind.NumberStoredAsText, ignored.Errors);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
