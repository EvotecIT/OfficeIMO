using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests;

public sealed class ExcelChartPhysicalLengthTests {
    [Fact]
    public void TransparentMarkerPaintSurvivesSnapshotAndSvgExportAfterReload() {
        using var bytes = new MemoryStream();
        using (var book = ExcelDocument.Create()) {
            var sheet = book.AddWorksheet("Markers");
            var chart = sheet.AddChart(new ExcelChartData(new[] { "First", "Next" },
                new[] { new ExcelChartSeries("Revenue", new[] { 12.5, 19.75 }) }), 1, 1,
                type: OfficeChartType.Line);
            chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle, OfficeLength.Points(8),
                fillColor: OfficeColor.Transparent, lineColor: OfficeColor.Transparent);
            book.Save(bytes);
        }
        bytes.Position = 0;
        using var loaded = ExcelDocument.Load(bytes);
        var retained = Assert.Single(loaded.GetSheet("Markers").Charts);
        Assert.True(retained.TryGetSnapshot(out var snapshot));
        var series = Assert.Single(snapshot.Data.Series);
        Assert.Equal("00000000", series.MarkerOutlineColorArgb);
        Assert.All(series.PointColorArgb!, color => Assert.Equal("00000000", color));
        var svg = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(retained.ToImage(
            new ExcelImageExportOptions { Format = OfficeIMO.Drawing.OfficeImageExportFormat.Svg })));
        var markers = svg.Descendants().Where(element => element.Name.LocalName is "ellipse" or "circle");
        Assert.All(markers, marker => {
            Assert.True(marker.Attribute("fill")?.Value == "none" || marker.Attribute("fill-opacity")?.Value == "0");
            Assert.True(marker.Attribute("stroke")?.Value is null or "none" || marker.Attribute("stroke-opacity")?.Value == "0");
        });
    }

    [Fact]
    public void ChartLengthsColorsAndPrintMarginsRetainTheirNativeUnitsAfterReload() {
        using var bytes = new MemoryStream();
        using (var book = ExcelDocument.Create()) {
            var sheet = book.AddWorksheet("Report");
            sheet.SetMargins(OfficeLength.Centimeters(2.54), OfficeLength.Inches(0.5),
                OfficeLength.Points(36), OfficeLength.Points(18));
            var chart = sheet.AddChart(new ExcelChartData(new[] { "First", "Next" },
                new[] { new ExcelChartSeries("Revenue", new[] { 12.5, 19.75 }) }), 2, 4,
                OfficeLength.Emus(4750001), OfficeLength.Points(150.25), type: OfficeChartType.Line, title: "Revenue");
            chart.SetTitleTextStyle(fontSize: OfficeLength.Points(12.345), color: OfficeColor.Accent1);
            chart.SetSeriesLineColor(0, OfficeColor.FromRgba(12, 34, 56, 128), OfficeLength.Emus(19051));
            chart.SetSeriesMarker(0, OfficeChartMarkerShape.Circle, OfficeLength.Points(8.5));
            Assert.True(chart.TryGetSnapshot(out var snapshot));
            Assert.Equal(4750001, snapshot.Width.ToEmus());
            Assert.Equal(1908175, snapshot.Height.ToEmus());
            book.Save(bytes);
        }
        bytes.Position = 0;
        using (var saved = SpreadsheetDocument.Open(bytes, false)) {
            var sheet = saved.WorkbookPart!.WorksheetParts.Single(part => part.DrawingsPart != null);
            var margins = sheet.Worksheet.GetFirstChild<PageMargins>()!;
            Assert.Equal(1d, margins.Left!.Value);
            Assert.Equal(0.5d, margins.Right!.Value);
            Assert.Equal(0.5d, margins.Top!.Value);
            Assert.Equal(0.25d, margins.Bottom!.Value);
            Assert.Equal(0.3d, margins.Header!.Value);
            Assert.Equal(0.3d, margins.Footer!.Value);
            var drawing = sheet.DrawingsPart!;
            var extent = Assert.Single(drawing.WorksheetDrawing!.Descendants<Xdr.Extent>());
            Assert.Equal(4750001, extent.Cx!.Value);
            Assert.Equal(1908175, extent.Cy!.Value);
            var chart = Assert.Single(drawing.ChartParts).ChartSpace;
            var title = Assert.Single(chart.Descendants<C.Title>());
            Assert.Contains(title.Descendants<A.RunProperties>(), props => props.FontSize?.Value == 1235);
            Assert.Contains(title.Descendants<A.SchemeColor>(), color => color.Val?.InnerText == "accent1");
            var series = Assert.Single(chart.Descendants<C.LineChartSeries>());
            Assert.Equal(9, series.GetFirstChild<C.Marker>()!.Size!.Val!.Value);
            var outline = Assert.Single(series.GetFirstChild<C.ChartShapeProperties>()!.Elements<A.Outline>());
            Assert.Equal(19051, outline.Width!.Value);
            Assert.Equal("0C2238", Assert.Single(outline.Descendants<A.RgbColorModelHex>()).Val!.Value);
            Assert.Equal(50196, Assert.Single(outline.Descendants<A.Alpha>()).Val!.Value);
        }
        bytes.Position = 0;
        using var loaded = ExcelDocument.Load(bytes);
        Assert.True(Assert.Single(loaded.GetSheet("Report").Charts).TryGetSnapshot(out var retained));
        Assert.Equal(4750001, retained.Width.ToEmus());
        Assert.Equal(1908175, retained.Height.ToEmus());
    }
}
