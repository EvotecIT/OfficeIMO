using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class ExcelChartBubbleSnapshotTests {
    /// <summary>
    /// The 480 px chart anchored at D8 is wider than a Letter page at 100 %, so it would
    /// continue on a second page. A wide page keeps each bubble painted once.
    /// </summary>
    private static ExcelPdfExportOptions WideBubblePage(ExcelPdfExportOptions? options = null) {
        options ??= new ExcelPdfExportOptions();
        options.PageSize = new PdfCore.PdfPageSize(OfficeLength.Points(900), OfficeLength.Points(700));
        return options;
    }

    [Theory]
    [InlineData(false, 100)]
    [InlineData(true, 100)]
    [InlineData(false, 200)]
    [InlineData(true, 200)]
    public void BubbleConversions_PaintSizedCirclesWithSourceRepresentationAndScale(bool width, int scale) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = CreateConversionBubble(document);
        C.BubbleChart native = GetNativeBubble(document);
        native.Append(new C.BubbleScale { Val = (uint)scale },
            new C.SizeRepresents { Val = width ? C.SizeRepresentsValues.Width : C.SizeRepresentsValues.Area });
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Equal(scale, snapshot.BubbleScalePercent);
        Assert.Equal(width ? OfficeChartBubbleSizeMode.Width : OfficeChartBubbleSizeMode.Area, snapshot.BubbleSizeMode);
        OfficeColor color = OfficeColor.Parse(snapshot.Data.Series[0].SeriesColorArgb!);
        double ratio = width ? 0.25D : 0.5D;

        var pdf = document.ConvertToPdfWithEvidence(WideBubblePage());
        Assert.DoesNotContain(pdf.Warnings, warning => warning.Code == "worksheet-chart" || warning.Code == "chart-approximation");
        double[] diameters = PdfBubbleDiameters(pdf.ToBytes(), color);
        Assert.Equal(2, diameters.Length);
        Assert.InRange(diameters[0] / diameters[1], ratio - 0.002D, ratio + 0.002D);

        var flow = document.ConvertToPdfWithEvidence(new ExcelPdfExportOptions { WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable });
        double[] flowDiameters = PdfBubbleDiameters(flow.ToBytes(), color);
        Assert.Equal(2, flowDiameters.Length);
        Assert.InRange(flowDiameters[0] / flowDiameters[1], ratio - 0.002D, ratio + 0.002D);

        var png = chart.ExportImage(OfficeImageExportFormat.Png);
        Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartKindApproximated);
        AssertRasterBubbleRatio(png.Bytes, color, ratio);
        // Exercise the worksheet PNG path as well as standalone chart export.
        AssertRasterBubbleRatio(document.Sheets[0].Range("A1:M28").ToImage(), color, ratio);

        // Compare scale to the same input at 100%, independently in PDF and raster output.
        if (scale == 200) {
            native.GetFirstChild<C.BubbleScale>()!.Val = 100U;
            double[] normal = PdfBubbleDiameters(document.ToPdf(WideBubblePage()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }), color);
            Assert.InRange(diameters[1] / normal[1], 1.999D, 2.001D);
            double[] normalRaster = RasterBubbleDiameters(chart.ExportImage(OfficeImageExportFormat.Png).Bytes, color);
            double[] enlargedRaster = RasterBubbleDiameters(png.Bytes, color);
            Assert.InRange(enlargedRaster[1] / normalRaster[1], 1.85D, 2.15D);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void BubbleConversions_NegativeSizesRetainSourceAndReportRequestedUnsupportedPainting(bool show, bool defaultAttribute) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = CreateConversionBubble(document);
        C.BubbleChart native = GetNativeBubble(document);
        native.Elements<C.BubbleChartSeries>().Single().GetFirstChild<C.BubbleSize>()!
            .Descendants<C.NumericPoint>().Single(point => point.Index!.Value == 0).NumericValue!.Text = "-4";
        native.Append(defaultAttribute ? new C.ShowNegativeBubbles() : new C.ShowNegativeBubbles { Val = show });
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Equal(-4D, snapshot.Data.Series[0].BubbleSizes![0]);
        Assert.Equal(show, snapshot.ShowNegativeBubbles);
        OfficeColor color = OfficeColor.Parse(snapshot.Data.Series[0].SeriesColorArgb!);
        var pdf = document.ConvertToPdfWithEvidence(WideBubblePage());
        Assert.Single(PdfBubbleDiameters(pdf.ToBytes(), color));
        Assert.Equal(show, pdf.Warnings.Any(warning => warning.Code == ExcelImageExportDiagnosticCodes.ChartBubbleNegativeSizesUnsupported));
        Assert.All(pdf.Warnings.Where(warning => warning.Code == ExcelImageExportDiagnosticCodes.ChartBubbleNegativeSizesUnsupported),
            warning => Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind));
        var png = chart.ExportImage(OfficeImageExportFormat.Png);
        Assert.Single(RasterBubbleDiameters(png.Bytes, color));
        Assert.Equal(show, png.Diagnostics.Any(diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartBubbleNegativeSizesUnsupported && diagnostic.LossKind == OfficeConversionLossKind.Omission));
    }

    [Fact]
    public void BubbleConversions_AutomaticSeriesColorsRemainDistinctInBothOutputs() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = CreateConversionBubble(document);
        var native = GetNativeBubble(document);
        var second = (C.BubbleChartSeries)native.Elements<C.BubbleChartSeries>().Single().CloneNode(true);
        second.GetFirstChild<C.Index>()!.Val = 1U;
        second.GetFirstChild<C.Order>()!.Val = 1U;
        foreach (var value in second.GetFirstChild<C.YValues>()!.Descendants<C.NumericValue>()) value.Text = "25";
        native.InsertAfter(second, native.Elements<C.BubbleChartSeries>().Single());
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        var colors = snapshot.Data.Series.Select(series => OfficeColor.Parse(series.SeriesColorArgb!)).ToArray();
        Assert.NotEqual(colors[0], colors[1]);
        byte[] pdf = document.ToPdf(WideBubblePage()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        byte[] png = chart.ExportImage(OfficeImageExportFormat.Png).Bytes;
        foreach (OfficeColor color in colors) {
            Assert.Equal(2, PdfBubbleDiameters(pdf, color).Length);
            Assert.Equal(2, RasterBubbleDiameters(png, color).Length);
        }
        var palette = new[] { OfficeColor.Parse("008080"), OfficeColor.Parse("A000A0") };
        byte[] overridden = document.ToPdf(WideBubblePage(new ExcelPdfExportOptions { ChartStyle = new OfficeChartStyle(palette: palette) })).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        foreach (OfficeColor color in palette) Assert.Equal(2, PdfBubbleDiameters(overridden, color).Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BubbleConversions_LeadingChartAnchorIsRetainedUnlessPrintAreaExcludesIt(bool printArea) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = CreateLiteralChart(document, false);
        var sheet = document.Sheets[0];
        sheet.CellValue(2, 2, "Source");
        sheet.CellValue(3, 3, 1);
        var marker = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing
            .Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker>().Single();
        marker.RowId!.Text = "0";
        marker.ColumnId!.Text = "0";
        if (printArea) document.SetPrintArea(sheet, "B2:C3");
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        OfficeColor color = OfficeColor.Parse(snapshot.Data.Series[0].SeriesColorArgb!);
        Assert.Equal(printArea ? 0 : 2, PdfBubbleDiameters(document.ToPdf(WideBubblePage()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }), color).Length);
    }

    [Fact]
    public void BubbleConversions_LeadingAnchorDoesNotMaterializeBlankFlowTableRows() {
        using var document = ExcelDocument.Create();
        CreateLiteralChart(document, false);
        document.Sheets[0].CellValue(1000, 1, "Far source");
        document.Sheets[0].CellValue(1001, 3, "Tail");
        byte[] bytes = document.ToPdf(new ExcelPdfExportOptions { WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var read = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Single(read.GetPages());
        Assert.Contains("Tail", read.GetPage(1).Text, StringComparison.Ordinal);
        // Canvas geometry must also retain those distant cells within the physical page.
        // At 100 % the chart and the distant cells print on their own pages; the blank
        // rows between them produce no pages.
        using var canvas = UglyToad.PdfPig.PdfDocument.Open(document.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.InRange(canvas.NumberOfPages, 1, 4);
        var page = Assert.Single(canvas.GetPages(), candidate => candidate.Text.Contains("Tail", StringComparison.Ordinal));
        Assert.All(page.Letters, letter => {
            Assert.InRange(letter.GlyphRectangle.Bottom, 0D, page.Height);
            Assert.InRange(letter.GlyphRectangle.Top, 0D, page.Height);
        });
    }

    [Fact]
    public void BubbleConversions_InvalidSettingsHaveDiagnosedDefaults() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = CreateConversionBubble(document);
        var native = GetNativeBubble(document);
        var scale = new C.BubbleScale(); scale.SetAttribute(new OpenXmlAttribute("val", string.Empty, "301"));
        var mode = new C.SizeRepresents(); mode.SetAttribute(new OpenXmlAttribute("val", string.Empty, "invalid"));
        native.Append(scale, mode);
        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Equal(100D, snapshot.BubbleScalePercent);
        Assert.Equal(OfficeChartBubbleSizeMode.Area, snapshot.BubbleSizeMode);
        Assert.Equal(2, chart.ExportImage(OfficeImageExportFormat.Png).Diagnostics.Count(item => item.Code == ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated));
        Assert.Equal(2, document.ConvertToPdfWithEvidence().Warnings.Count(item => item.Code == ExcelImageExportDiagnosticCodes.ChartBubbleSettingsApproximated));
    }

    private static ExcelChart CreateConversionBubble(ExcelDocument document) {
        var chart = CreateLiteralChart(document, false);
        document.Sheets[0].CellValue(1, 1, "Bubble conversion");
        var native = GetNativeBubble(document);
        native.RemoveAllChildren<C.BubbleScale>();
        native.RemoveAllChildren<C.SizeRepresents>();
        native.RemoveAllChildren<C.ShowNegativeBubbles>();
        document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.ChartParts.Single()
            .ChartSpace.GetFirstChild<C.Chart>()!.RemoveAllChildren<C.Legend>();
        return chart;
    }

    private static double[] PdfBubbleDiameters(byte[] bytes, OfficeColor color) =>
        PdfCore.PdfDocument.Load(new System.IO.MemoryStream(bytes)).Pages.SelectMany(page => page.GetIdentityVisualPrimitives())
            .Where(primitive => primitive.Kind == PdfCore.PdfPageVisualPrimitiveKind.Path && primitive.FillColor == color &&
                primitive.Width > 2D && Math.Abs(primitive.Width - primitive.Height) < 0.05D)
            .Select(primitive => primitive.Width).OrderBy(value => value).ToArray();

    private static void AssertRasterBubbleRatio(byte[] bytes, OfficeColor color, double ratio) {
        double[] sizes = RasterBubbleDiameters(bytes, color);
        Assert.Equal(2, sizes.Length);
        Assert.InRange(sizes[0] / sizes[1], ratio - 0.08D, ratio + 0.08D);
    }

    private static double[] RasterBubbleDiameters(byte[] bytes, OfficeColor color) {
        Assert.True(OfficePngReader.TryDecode(bytes, out var image));
        var seen = new HashSet<(int X, int Y)>();
        var sizes = new List<double>();
        bool Matches(int x, int y) => x >= 0 && x < image!.Width && y >= 0 && y < image.Height && image.GetPixel(x, y) == color;
        for (int y = 0; y < image!.Height; y++) for (int x = 0; x < image.Width; x++) {
            if (!Matches(x, y) || !seen.Add((x, y))) continue;
            var queue = new Queue<(int X, int Y)>();queue.Enqueue((x, y));
            int left = x, right = x, top = y, bottom = y;
            while (queue.Count > 0) {
                var p = queue.Dequeue();left = Math.Min(left, p.X);right = Math.Max(right, p.X);top = Math.Min(top, p.Y);bottom = Math.Max(bottom, p.Y);
                foreach (var next in new[] { (X: p.X - 1, Y: p.Y), (X: p.X + 1, Y: p.Y), (X: p.X, Y: p.Y - 1), (X: p.X, Y: p.Y + 1) })
                    if (Matches(next.X, next.Y) && seen.Add(next)) queue.Enqueue(next);
            }
            int width = right - left + 1, height = bottom - top + 1;
            if (width > 2 && height > 2) { Assert.InRange(Math.Abs(width - height), 0, 2);sizes.Add((width + height) / 2D); }
        }
        return sizes.OrderBy(value => value).ToArray();
    }
}
