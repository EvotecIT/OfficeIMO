using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetChartPointStyleConversionTests {
    [Theory]
    [InlineData(OdsChartType.Pie, ExcelChartType.Pie)]
    [InlineData(OdsChartType.Doughnut, ExcelChartType.Doughnut)]
    public void OdsRadialPointStylesConvertToEditableExcelCharts(OdsChartType sourceType,
        ExcelChartType targetType) {
        OdsDocument source = CreateStyledOdsChart(sourceType);

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        ExcelChart chart = Assert.Single(converted["Data"].Charts);
        Assert.Equal(targetType, chart.ChartType);
        Assert.DoesNotContain(result.Report.Mappings, mapping =>
            mapping.Feature == "source-embedded-objects" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);

        using (SpreadsheetDocument package = SpreadsheetDocument.Open(
            new MemoryStream(converted.ToBytes()), false)) {
            ChartPart part = Assert.Single(package.WorkbookPart!.WorksheetParts
                .SelectMany(sheet => sheet.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>()));
            C.DataPoint[] points = part.ChartSpace!.Descendants<C.DataPoint>().ToArray();
            Assert.Equal(3, points.Length);
            Assert.Equal(new uint[] { 0, 1, 2 }, points.Select(point => point.Index!.Val!.Value));
            Assert.NotNull(points[0].GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.SolidFill>());
            Assert.NotNull(points[1].GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.PatternFill>());
            C.ChartShapeProperties outlined = points[2].GetFirstChild<C.ChartShapeProperties>()!;
            Assert.NotNull(outlined.GetFirstChild<A.NoFill>());
            Assert.NotNull(outlined.GetFirstChild<A.Outline>());
        }

        OdsDocument roundTripped = converted.ToOpenDocumentResult().Value;
        OfficeChartPointStyle?[] styles = Assert.Single(
            Assert.Single(roundTripped.GetSheet("Data")!.Charts).Series).PointStyles!.ToArray();
        Assert.Equal(OfficeColor.Parse("#228844"), styles[0]!.FillColor);
        Assert.Equal(OfficeChartHatchPattern.ForwardDiagonal, styles[1]!.Hatch);
        Assert.True(styles[2]!.NoFill);
        OdfValidationResult validation = OdsDocument.Load(new MemoryStream(roundTripped.ToBytes())).Validate();
        Assert.True(validation.IsValid, string.Join("; ", validation.Diagnostics.Select(item =>
            item.Id + ": " + item.Message)));
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie, "chart:circle")]
    [InlineData(OfficeChartKind.Doughnut, "chart:ring")]
    public void ExcelSharedPointStylesConvertToNativeOds(OfficeChartKind sourceType,
        string chartClass) {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        var data = new OfficeChartData(new[] { "Pass", "Fail", "Unknown" },
            new[] { new OfficeChartSeries("Status", new[] { 3d, 4d, 5d })
                .WithPointStyles(CreatePointStyles()) });
        sheet.AddChart(sourceType, data, row: 2, column: 4, title: "Status");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        OdsChart chart = Assert.Single(result.Value.GetSheet("Summary")!.Charts);
        Assert.Equal(chartClass, chart.ChartClass);
        OfficeChartPointStyle?[] styles = Assert.Single(chart.Series).PointStyles!.ToArray();
        Assert.Equal(OfficeColor.Parse("#228844"), styles[0]!.FillColor);
        Assert.Equal(OfficeChartHatchPattern.ForwardDiagonal, styles[1]!.Hatch);
        Assert.Equal(OfficeColor.Parse("#D97706"), styles[1]!.HatchColor);
        Assert.True(styles[2]!.NoFill);
        Assert.Equal(OfficeColor.Parse("#445566"), styles[2]!.OutlineColor);
        Assert.DoesNotContain(result.Report.Mappings, mapping =>
            mapping.Feature == "charts" && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.True(OdsDocument.Load(new MemoryStream(result.Value.ToBytes())).Validate().IsValid);
    }

    [Fact]
    public void UnsupportedOdsPointAppearanceReportsTheWholeChartAsLoss() {
        OdsDocument source = CreateStyledOdsChart(OdsChartType.Pie);
        XDocument styles = XDocument.Parse(Encoding.UTF8.GetString(
            source.GetPackageEntryBytes("Object 1/styles.xml")));
        Assert.Single(styles.Descendants(OdfNamespaces.Draw + "hatch"))
            .SetAttributeValue(OdfNamespaces.Draw + "distance", "0.8cm");
        source.Package.AddOrReplaceEntry("Object 1/styles.xml",
            Encoding.UTF8.GetBytes(styles.ToString(SaveOptions.DisableFormatting)), "text/xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "source-embedded-objects" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void OdsImplicitTrailingPointsKeepTheStyledChartConvertible() {
        OdsDocument source = CreateStyledOdsChart(OdsChartType.Pie);
        XDocument part = XDocument.Parse(Encoding.UTF8.GetString(
            source.GetPackageEntryBytes("Object 1/content.xml")));
        XElement series = Assert.Single(part.Descendants(OdfNamespaces.Chart + "series"));
        foreach (XElement point in series.Elements(OdfNamespaces.Chart + "data-point").Skip(1).ToArray())
            point.Remove();
        source.Package.AddOrReplaceEntry("Object 1/content.xml",
            Encoding.UTF8.GetBytes(part.ToString(SaveOptions.DisableFormatting)), "text/xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Single(converted["Data"].Charts);
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(
            new MemoryStream(converted.ToBytes()), false)) {
            ChartPart chartPart = Assert.Single(package.WorkbookPart!.WorksheetParts
                .SelectMany(sheet => sheet.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>()));
            C.DataPoint point = Assert.Single(chartPart.ChartSpace!.Descendants<C.DataPoint>());
            Assert.Equal(0U, point.Index!.Val!.Value);
        }
        Assert.DoesNotContain(result.Report.Mappings, mapping =>
            mapping.Feature == "source-embedded-objects" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void PartiallySupportedOdsPointEffectReportsWholeChartLoss() {
        OdsDocument source = CreateStyledOdsChart(OdsChartType.Pie);
        XDocument part = XDocument.Parse(Encoding.UTF8.GetString(
            source.GetPackageEntryBytes("Object 1/content.xml")));
        XElement point = Assert.Single(part.Descendants(OdfNamespaces.Chart + "series"))
            .Element(OdfNamespaces.Chart + "data-point")!;
        string name = (string)point.Attribute(OdfNamespaces.Chart + "style-name")!;
        XElement definition = Assert.Single(part.Descendants(OdfNamespaces.Style + "style"),
            item => (string?)item.Attribute(OdfNamespaces.Style + "name") == name);
        definition.Element(OdfNamespaces.Style + "graphic-properties")!
            .SetAttributeValue(OdfNamespaces.Draw + "shadow", "visible");
        source.Package.AddOrReplaceEntry("Object 1/content.xml",
            Encoding.UTF8.GetBytes(part.ToString(SaveOptions.DisableFormatting)), "text/xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "source-embedded-objects" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void StyledOdsLineReportsWholeChartLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        for (int index = 0; index < 3; index++) {
            sheet.Cell(index, 0).SetString("Category " + index);
            sheet.Cell(index, 1).SetNumber(index + 1);
        }
        sheet.AddChart(OdsChartType.Line, "Data.$A$1:.$A$3",
            new[] { new OdsChartSeries("Data.$B$1:.$B$3") },
            2, 4, OdfRect.FromCentimeters(0, 0, 10, 7));
        XDocument part = XDocument.Parse(Encoding.UTF8.GetString(
            source.GetPackageEntryBytes("Object 1/content.xml")));
        XElement series = Assert.Single(part.Descendants(OdfNamespaces.Chart + "series"));
        XElement point = Assert.Single(series.Elements(OdfNamespaces.Chart + "data-point"));
        point.SetAttributeValue(OdfNamespaces.Chart + "repeated", null);
        point.SetAttributeValue(OdfNamespaces.Chart + "style-name", "StyledLinePoint");
        part.Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(
            new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "name", "StyledLinePoint"),
                new XAttribute(OdfNamespaces.Style + "family", "chart"),
                new XElement(OdfNamespaces.Style + "graphic-properties",
                    new XAttribute(OdfNamespaces.Draw + "fill", "solid"),
                    new XAttribute(OdfNamespaces.Draw + "fill-color", "#228844"))));
        source.Package.AddOrReplaceEntry("Object 1/content.xml",
            Encoding.UTF8.GetBytes(part.ToString(SaveOptions.DisableFormatting)), "text/xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "source-embedded-objects" &&
            mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void StyledExcelLineReportsWholeChartLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        var data = new OfficeChartData(new[] { "Pass", "Fail", "Unknown" },
            new[] { new OfficeChartSeries("Status", new[] { 3d, 4d, 5d })
                .WithPointStyles(new OfficeChartPointStyle?[] {
                    new(fillColor: OfficeColor.Parse("#228844")), null, null
                }) });
        sheet.AddChart(OfficeChartKind.Line, data, row: 2, column: 4);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "charts" && mapping.Status == OdfConversionMappingStatus.Unsupported &&
            mapping.Count == 1);
    }

    [Fact]
    public void OdsDoughnutKeepsStylesOnBothRingsThroughExcel() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        for (int index = 0; index < 3; index++) {
            sheet.Cell(index, 0).SetString(new[] { "Pass", "Fail", "Unknown" }[index]);
            sheet.Cell(index, 1).SetNumber(index + 3);
            sheet.Cell(index, 2).SetNumber(index + 6);
        }
        sheet.AddChart(OdsChartType.Doughnut, "Data.$A$1:.$A$3", new[] {
            new OdsChartSeries("Data.$B$1:.$B$3").WithPointStyles(CreatePointStyles()),
            new OdsChartSeries("Data.$C$1:.$C$3").WithPointStyles(new OfficeChartPointStyle?[] {
                new(fillColor: OfficeColor.Parse("#4466AA")), null, null
            })
        }, 2, 4, OdfRect.FromCentimeters(0, 0, 10, 7));

        using ExcelDocument excel = source.ToExcelDocumentResult().Value;
        ExcelChart converted = Assert.Single(excel["Data"].Charts);
        Assert.True(converted.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(2, snapshot.Data.Series.Count);
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(excel.ToBytes()), false)) {
            ChartPart part = Assert.Single(package.WorkbookPart!.WorksheetParts
                .SelectMany(item => item.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>()));
            C.PieChartSeries[] rings = part.ChartSpace!.Descendants<C.PieChartSeries>().ToArray();
            Assert.Equal(2, rings.Length);
            Assert.Equal(new[] { 3, 1 }, rings.Select(ring => ring.Elements<C.DataPoint>().Count()));
            Assert.NotNull(Assert.Single(rings[1].Elements<C.DataPoint>())
                .GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.SolidFill>());
        }

        OdsDocument restored = excel.ToOpenDocumentResult().Value;
        OdsChart chart = Assert.Single(restored.GetSheet("Data")!.Charts);
        Assert.Equal(2, chart.Series.Count);
        Assert.Equal(OfficeColor.Parse("#4466AA"), chart.Series[1].PointStyles![0]!.FillColor);
        Assert.True(OdsDocument.Load(new MemoryStream(restored.ToBytes())).Validate().IsValid);
    }

    [Fact]
    public void ExplodedExcelSliceReportsTheWholeChartAsLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        var data = new OfficeChartData(new[] { "Pass", "Fail", "Unknown" },
            new[] { new OfficeChartSeries("Status", new[] { 3d, 4d, 5d })
                .WithPointStyles(CreatePointStyles()).WithPointExplosions(new[] { 0, 25, 0 }) });
        sheet.AddChart(OfficeChartKind.Pie, data, row: 2, column: 4);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Empty(result.Value.GetSheet("Summary")!.Charts);
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "charts" && mapping.Status == OdfConversionMappingStatus.Unsupported &&
            mapping.Count == 1);
    }

    [Fact]
    public void TransparentExcelPointReportsChartLossWithoutAbortingWorkbook() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Summary");
        sheet.Cell(1, 1, "Retained cell");
        var data = new OfficeChartData(new[] { "Pass", "Fail", "Unknown" },
            new[] { new OfficeChartSeries("Status", new[] { 3d, 4d, 5d })
                .WithPointStyles(new OfficeChartPointStyle?[] {
                    new(fillColor: OfficeColor.FromRgba(34, 136, 68, 128)), null, null
                }) });
        sheet.AddChart(OfficeChartKind.Pie, data, row: 2, column: 4);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        OdsSheet converted = result.Value.GetSheet("Summary")!;
        Assert.Equal("Retained cell", converted.Cell(0, 0).Value.DisplayText);
        Assert.Empty(converted.Charts);
        Assert.Contains(result.Report.Mappings, mapping =>
            mapping.Feature == "charts" && mapping.Status == OdfConversionMappingStatus.Unsupported &&
            mapping.Count == 1);
    }

    private static OdsDocument CreateStyledOdsChart(OdsChartType type) {
        OdsDocument document = OdsDocument.Create();
        OdsSheet sheet = document.AddSheet("Data");
        for (int index = 0; index < 3; index++) {
            sheet.Cell(index, 0).SetString(new[] { "Pass", "Fail", "Unknown" }[index]);
            sheet.Cell(index, 1).SetNumber(index + 3);
        }
        sheet.AddChart(type, "Data.$A$1:.$A$3",
            new[] { new OdsChartSeries("Data.$B$1:.$B$3")
                .WithPointStyles(CreatePointStyles()) },
            2, 4, OdfRect.FromCentimeters(0, 0, 10, 7), "Status");
        return document;
    }

    private static OfficeChartPointStyle?[] CreatePointStyles() => new OfficeChartPointStyle?[] {
        new(fillColor: OfficeColor.Parse("#228844")),
        new(fillColor: OfficeColor.Parse("#FFF0DD"), hatch: OfficeChartHatchPattern.ForwardDiagonal,
            hatchColor: OfficeColor.Parse("#D97706")),
        new(noFill: true, outlineColor: OfficeColor.Parse("#445566"),
            outlineWidth: 2, showOutline: true)
    };
}
