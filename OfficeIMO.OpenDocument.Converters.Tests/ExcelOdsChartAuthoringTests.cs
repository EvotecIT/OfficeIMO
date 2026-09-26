using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.Spreadsheet;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class ExcelOdsChartAuthoringTests {
    [Theory]
    [InlineData(ExcelChartType.ColumnClustered, "chart:bar", false)]
    [InlineData(ExcelChartType.BarClustered, "chart:bar", true)]
    [InlineData(ExcelChartType.Line, "chart:line", null)]
    public void ExcelChartConvertsToReopenableOdsAndBack(ExcelChartType type, string chartClass, bool? vertical) {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        var data = new ExcelChartData(new[] { "Jan", "Feb" },
            new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) });
        sheet.AddChart(data, row: 5, column: 4, widthPixels: 480, heightPixels: 300,
            type: type, title: "Sales");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
            mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
        OdsChart chart = Assert.Single(result.Value.GetSheet("Summary")!.Charts);
        SpreadsheetRangeReference categories = SpreadsheetRangeReference.Parse(
            chart.CategoriesAddress!, SpreadsheetAddressDialect.OpenDocument);
        string sourceSheetName = categories.Start.SheetName!;
        Assert.NotEqual("Summary", sourceSheetName);
        Assert.NotNull(result.Value.GetSheet(sourceSheetName));
        Assert.Equal(sourceSheetName, categories.End!.SheetName);
        SpreadsheetRangeReference values = SpreadsheetRangeReference.Parse(
            Assert.Single(chart.Series).ValuesAddress, SpreadsheetAddressDialect.OpenDocument);
        Assert.Equal(sourceSheetName, values.End!.SheetName);
        Assert.Equal(chartClass, chart.ChartClass);
        Assert.Equal(vertical, chart.VerticalBars);
        Assert.Equal("Sales", chart.Title);

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(result.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        Assert.Single(reopened.GetSheet("Summary")!.Charts);
        OdfConversionResult<ExcelDocument> back = reopened.ToExcelDocumentResult();
        using ExcelDocument roundTrip = back.Value;
        ExcelChart converted = Assert.Single(roundTrip["Summary"].Charts);
        Assert.Equal(type, converted.ChartType);
        Assert.True(converted.TryGetData(out ExcelChartData actual));
        Assert.Equal(new[] { "Jan", "Feb" }, actual.Categories);
        Assert.Equal(new[] { 10d, 20d }, Assert.Single(actual.Series).Values);
        Assert.Throws<OdfConversionLossException>(() => source.ToOpenDocumentResult(
            new ExcelOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void UnsupportedExcelPieChartRemainsExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        sheet.AddChart(new ExcelChartData(new[] { "Jan", "Feb" },
            new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) }),
            row: 5, column: 4, type: ExcelChartType.Pie);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void AbsoluteAnchoredChartRemainsExplicitLoss() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-chart-absolute-" +
            Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument source = ExcelDocument.Create(path)) {
                ExcelSheet sheet = source.AddWorksheet("Summary");
                sheet.AddChart(new ExcelChartData(new[] { "Jan", "Feb" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) }),
                    row: 5, column: 4, type: ExcelChartType.ColumnClustered);
                source.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                var drawing = package.WorkbookPart!.WorksheetParts
                    .Select(part => part.DrawingsPart?.WorksheetDrawing)
                    .Single(root => root != null)!;
                Xdr.OneCellAnchor anchor = drawing.Elements<Xdr.OneCellAnchor>().Single();
                Xdr.GraphicFrame frame = anchor.GetFirstChild<Xdr.GraphicFrame>()!;
                anchor.InsertAfterSelf(new Xdr.AbsoluteAnchor(
                    new Xdr.Position { X = 4572000L, Y = 914400L },
                    new Xdr.Extent { Cx = 4572000L, Cy = 2743200L },
                    (Xdr.GraphicFrame)frame.CloneNode(true), new Xdr.ClientData()));
                anchor.Remove();
                drawing.Save();
            }
            using ExcelDocument imported = ExcelDocument.Load(path);
            OdfConversionResult<OdsDocument> result = imported.ToOpenDocumentResult();
            Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
                mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void BrokenChartRelationshipRemainsExplicitLoss() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-chart-broken-" +
            Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument source = ExcelDocument.Create(path)) {
                source.AddWorksheet("Summary").AddChart(new ExcelChartData(new[] { "Jan" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d }) }), row: 5, column: 4);
                source.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Xdr.GraphicFrame frame = package.WorkbookPart!.WorksheetParts
                    .Select(part => part.DrawingsPart?.WorksheetDrawing)
                    .Single(root => root != null)!.Descendants<Xdr.GraphicFrame>().Single();
                frame.Graphic!.GraphicData!.GetFirstChild<C.ChartReference>()!.Id = "rIdMissing";
                frame.Ancestors<Xdr.WorksheetDrawing>().Single().Save();
            }
            using ExcelDocument imported = ExcelDocument.Load(path);
            OdfConversionResult<OdsDocument> result = imported.ToOpenDocumentResult();
            Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
                mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SharedChartPartFramesAreCountedSeparately() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-chart-shared-" +
            Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument source = ExcelDocument.Create(path)) {
                source.AddWorksheet("Summary").AddChart(new ExcelChartData(new[] { "Jan", "Feb" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) }), row: 5, column: 4);
                source.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                Xdr.WorksheetDrawing drawing = package.WorkbookPart!.WorksheetParts
                    .Select(part => part.DrawingsPart?.WorksheetDrawing)
                    .Single(root => root != null)!;
                Xdr.GraphicFrame frame = drawing.Descendants<Xdr.GraphicFrame>().Single();
                drawing.Append(new Xdr.AbsoluteAnchor(new Xdr.Position { X = 4572000L, Y = 914400L },
                    new Xdr.Extent { Cx = 4572000L, Cy = 2743200L },
                    (Xdr.GraphicFrame)frame.CloneNode(true), new Xdr.ClientData()));
                drawing.Save();
            }
            using ExcelDocument imported = ExcelDocument.Load(path);
            OdfConversionResult<OdsDocument> result = imported.ToOpenDocumentResult();
            Assert.Single(result.Value.GetSheet("Summary")!.Charts);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
                mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
                mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void SingleCategoryHorizontalChartRoundTripsThroughOds() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-chart-horizontal-" +
            Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument source = ExcelDocument.Create(path)) {
                ExcelSheet sheet = source.AddWorksheet("Summary");
                var data = new ExcelChartData(new[] { "Jan" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d }) });
                ExcelChartDataRange range = sheet.WriteChartData(data, orientation: ExcelChartDataOrientation.Horizontal);
                sheet.AddChart(range, row: 5, column: 4, type: ExcelChartType.ColumnClustered);
                source.Save();
            }
            using ExcelDocument imported = ExcelDocument.Load(path);
            OdfConversionResult<OdsDocument> result = imported.ToOpenDocumentResult();
            Assert.Single(result.Value.GetSheet("Summary")!.Charts);
            Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
                mapping.Status == OdfConversionMappingStatus.Unsupported);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void MixedSeriesAndSecondaryAxisRemainExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        sheet.AddChart(new ExcelChartData(new[] { "Jan", "Feb" }, new[] {
            new ExcelChartSeries("Sales", new[] { 10d, 20d }, ExcelChartType.ColumnClustered),
            new ExcelChartSeries("Trend", new[] { 12d, 22d }, ExcelChartType.Line,
                OfficeChartAxisGroup.Secondary)
        }), row: 5, column: 4, type: ExcelChartType.ColumnClustered);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void NonnumericWorksheetValueRemainsExplicitChartLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        ExcelChart chart = sheet.AddChart(new ExcelChartData(new[] { "Jan", "Feb" },
            new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) }),
            row: 5, column: 4, type: ExcelChartType.ColumnClustered);
        ExcelChartDataRange range = chart.DataRange!;
        source[range.SheetName].Cell(range.SeriesStartRow, range.SeriesStartColumn, "not numeric");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DivergentSeriesCategoryOrLabelReferencesRemainExplicitLoss(bool changeLabel) {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-chart-references-" + Guid.NewGuid().ToString("N") + ".xlsx");
        string dataSheet;
        try {
            using (ExcelDocument source = ExcelDocument.Create(path)) {
                ExcelSheet sheet = source.AddWorksheet("Summary");
                ExcelChart chart = sheet.AddChart(new ExcelChartData(new[] { "Jan", "Feb" }, new[] {
                    new ExcelChartSeries("Sales", new[] { 10d, 20d }),
                    new ExcelChartSeries("Trend", new[] { 12d, 22d })
                }), row: 5, column: 4, type: ExcelChartType.ColumnClustered);
                dataSheet = chart.DataRange!.SheetName;
                source.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                C.ChartSpace space = package.WorkbookPart!.WorksheetParts
                    .SelectMany(part => part.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
                    .Single().ChartSpace!;
                C.BarChartSeries second = space.Descendants<C.BarChartSeries>().Skip(1).Single();
                if (changeLabel)
                    second.GetFirstChild<C.SeriesText>()!.GetFirstChild<C.StringReference>()!
                        .Formula!.Text = dataSheet + "!$B$2";
                else
                    second.GetFirstChild<C.CategoryAxisData>()!.GetFirstChild<C.StringReference>()!
                        .Formula!.Text = dataSheet + "!$A$3:$A$4";
                space.Save();
            }
            using ExcelDocument reloaded = ExcelDocument.Load(path);
            OdfConversionResult<OdsDocument> result = reloaded.ToOpenDocumentResult();
            Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts" &&
                mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void CaseVariantExcelSheetReferencesUseCanonicalOdsSheetName() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-chart-case-" + Guid.NewGuid().ToString("N") + ".xlsx");
        string dataSheet;
        try {
            using (ExcelDocument source = ExcelDocument.Create(path)) {
                ExcelSheet sheet = source.AddWorksheet("Summary");
                ExcelChart sourceChart = sheet.AddChart(new ExcelChartData(new[] { "Jan", "Feb" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) }),
                    row: 5, column: 4, type: ExcelChartType.ColumnClustered);
                dataSheet = sourceChart.DataRange!.SheetName;
                source.Save();
            }
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                C.ChartSpace space = package.WorkbookPart!.WorksheetParts
                    .SelectMany(part => part.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>())
                    .Single().ChartSpace!;
                foreach (C.Formula formula in space.Descendants<C.Formula>())
                    if (formula.Text is string text)
                        formula.Text = text.Replace(dataSheet, dataSheet.ToLowerInvariant());
                space.Save();
            }
            using ExcelDocument reloaded = ExcelDocument.Load(path);
            OdsChart chart = Assert.Single(reloaded.ToOpenDocumentResult().Value.GetSheet("Summary")!.Charts);
            SpreadsheetRangeReference address = SpreadsheetRangeReference.Parse(
                chart.CategoriesAddress!, SpreadsheetAddressDialect.OpenDocument);
            Assert.Equal(dataSheet, address.Start.SheetName);
            Assert.Equal(dataSheet, address.End!.SheetName);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
