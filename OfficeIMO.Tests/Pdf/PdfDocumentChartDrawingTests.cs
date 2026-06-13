using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentChartDrawingTests {
    [Fact]
    public void FlowDrawing_RenderWithQuality_ReturnsSharedQualityReport() {
        OfficeChartRenderingResult result = OfficeChartDrawingRenderer.RenderWithQuality(new OfficeChartSnapshot(
            "Quality chart",
            "Quality Report",
            OfficeChartKind.Line,
            new OfficeChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 12D, 18D, 24D })
                }),
            widthPoints: 320D,
            heightPoints: 190D));

        Assert.NotNull(result.Drawing);
        Assert.NotNull(result.QualityReport);
        Assert.False(result.QualityReport.HasIssues, string.Join("; ", result.QualityReport.Issues.Select(issue => issue.ToString())));
    }

    [Fact]
    public void ScatterRange_IncludesSharedXValuesWhenSeriesMixExplicitAndSharedCoordinates() {
        var series = new[] {
            new OfficeChartSeries("Explicit", new[] { 4D, 5D }, new[] { 10D, 20D }),
            new OfficeChartSeries("Shared", new[] { 1D, 2D })
        };
        MethodInfo method = typeof(OfficeChartDrawingRenderer).GetMethod("GetScatterXRange", BindingFlags.NonPublic | BindingFlags.Static)!;

        object range = method.Invoke(null, new object[] { series, new[] { 100D, 200D } })!;

        double min = (double)range.GetType().GetProperty("Min")!.GetValue(range)!;
        double max = (double)range.GetType().GetProperty("Max")!.GetValue(range)!;
        Assert.Equal(10D, min);
        Assert.Equal(200D, max);
    }

    [Fact]
    public void PercentStackedRange_PreservesNegativeStacksBelowZero() {
        var series = new[] {
            new OfficeChartSeries("Positive", new[] { 6D, 4D }),
            new OfficeChartSeries("Negative", new[] { -3D, 2D })
        };
        MethodInfo method = typeof(OfficeChartDrawingRenderer).GetMethod("GetPercentStackedSeriesRange", BindingFlags.NonPublic | BindingFlags.Static)!;

        object range = method.Invoke(null, new object[] { series, 2 })!;

        double min = (double)range.GetType().GetProperty("Min")!.GetValue(range)!;
        double max = (double)range.GetType().GetProperty("Max")!.GetValue(range)!;
        Assert.Equal(-1D, min);
        Assert.Equal(1D, max);
    }

    [Fact]
    public void FlowDrawing_RendersScatterXAxisLabelsFromNumericValues() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Scatter",
            "Scatter Axis",
            OfficeChartKind.Scatter,
            new OfficeChartData(
                new[] { "Alpha", "Beta", "Gamma" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 3D, 4D, 5D }, new[] { 1D, 10D, 100D })
                }),
            widthPoints: 320D,
            heightPoints: 190D));

        var labels = drawing.Elements.OfType<OfficeDrawingText>().Select(text => text.Text).ToList();

        Assert.Contains("1", labels);
        Assert.Contains("100", labels);
        Assert.DoesNotContain("Alpha", labels);
        Assert.DoesNotContain("Beta", labels);
        Assert.DoesNotContain("Gamma", labels);
    }

    [Fact]
    public void FlowDrawing_RendersSinglePointLineChartMarker() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Single point",
            "Line Marker",
            OfficeChartKind.Line,
            new OfficeChartData(
                new[] { "Only" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 42D })
                }),
            widthPoints: 220D,
            heightPoints: 140D));

        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Ellipse &&
            shape.Shape.Width == 4D &&
            shape.Shape.Height == 4D);
    }

    [Fact]
    public void FlowDrawing_RendersSingleCategoryBarChart() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Single bar",
            "One Category",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Only" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 42D })
                }),
            widthPoints: 220D,
            heightPoints: 140D));

        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Rectangle &&
            shape.Shape.Width > 2D &&
            shape.Shape.Height > 2D &&
            shape.Shape.FillColor.HasValue &&
            shape.Shape.StrokeWidth == 0D);
    }

    [Fact]
    public void FlowDrawing_RendersBarChartPointColors() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Point colors",
            "Per Point",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new OfficeChartSeries(
                        "Actual",
                        new[] { 10D, 20D },
                        null,
                        OfficeColor.Black,
                        new OfficeColor?[] { OfficeColor.ParseHex("#2FB344"), OfficeColor.ParseHex("#F76707") })
                }),
            widthPoints: 220D,
            heightPoints: 140D));

        var barColors = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Rectangle && shape.Shape.StrokeWidth == 0D)
            .Select(shape => shape.Shape.FillColor)
            .ToList();

        Assert.Contains(OfficeColor.ParseHex("#2FB344"), barColors);
        Assert.Contains(OfficeColor.ParseHex("#F76707"), barColors);
    }

    [Fact]
    public void FlowDrawing_RendersPieChartSeriesColorWhenPointColorsAreMissing() {
        OfficeColor seriesColor = OfficeColor.ParseHex("#CC3366");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Pie series color",
            "Pie Series",
            OfficeChartKind.Pie,
            new OfficeChartData(
                new[] { "Passed", "Failed" },
                new[] {
                    new OfficeChartSeries(
                        "Outcome",
                        new[] { 42D, 30D },
                        null,
                        seriesColor)
                }),
            widthPoints: 260D,
            heightPoints: 180D));

        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Polygon &&
            shape.Shape.FillColor == seriesColor);
        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Rectangle &&
            shape.Shape.FillColor == seriesColor &&
            shape.Shape.StrokeWidth == 0D);
    }

    [Fact]
    public void FlowDrawing_RendersDoughnutChartSeriesColorWhenPointColorsAreMissing() {
        OfficeColor seriesColor = OfficeColor.ParseHex("#CC3366");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Doughnut series color",
            "Doughnut Series",
            OfficeChartKind.Doughnut,
            new OfficeChartData(
                new[] { "Passed", "Failed" },
                new[] {
                    new OfficeChartSeries(
                        "Outcome",
                        new[] { 42D, 30D },
                        null,
                        seriesColor)
                }),
            widthPoints: 260D,
            heightPoints: 180D));

        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Polygon &&
            shape.Shape.FillColor == seriesColor);
        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Rectangle &&
            shape.Shape.FillColor == seriesColor &&
            shape.Shape.StrokeWidth == 0D);
    }

    [Fact]
    public void FlowDrawing_RendersLineChartPointColorsOnMarkers() {
        OfficeColor highlight = OfficeColor.ParseHex("#F76707");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Line point colors",
            "Line Points",
            OfficeChartKind.Line,
            new OfficeChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeChartSeries(
                        "Actual",
                        new[] { 10D, 20D, 14D },
                        null,
                        OfficeColor.ParseHex("#2563EB"),
                        new OfficeColor?[] { null, highlight, null })
                }),
            widthPoints: 260D,
            heightPoints: 160D));

        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Ellipse &&
            shape.Shape.Width == 4D &&
            shape.Shape.Height == 4D &&
            shape.Shape.FillColor == highlight &&
            shape.Shape.StrokeColor == highlight);
    }

    [Fact]
    public void FlowDrawing_SkipsNonFiniteScatterXCoordinates() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Scatter",
            "Finite Points",
            OfficeChartKind.Scatter,
            new OfficeChartData(
                new[] { "1", "2", "3" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 3D, 4D, 5D }, new[] { 1D, double.NaN, 3D })
                }),
            widthPoints: 320D,
            heightPoints: 190D));

        int markerCount = drawing.Shapes.Count(shape =>
            shape.Shape.Kind == OfficeShapeKind.Ellipse &&
            shape.Shape.Width == 5D &&
            shape.Shape.Height == 5D);
        Assert.Equal(2, markerCount);
    }

    [Fact]
    public void FlowDrawing_RendersScatterPointColorsUsingSourcePointIndex() {
        OfficeColor highlight = OfficeColor.ParseHex("#2FB344");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Scatter point colors",
            "Scatter Points",
            OfficeChartKind.Scatter,
            new OfficeChartData(
                new[] { "1", "2", "3" },
                new[] {
                    new OfficeChartSeries(
                        "Actual",
                        new[] { 3D, 4D, 5D },
                        new[] { 1D, double.NaN, 3D },
                        OfficeColor.ParseHex("#2563EB"),
                        new OfficeColor?[] { null, highlight, OfficeColor.ParseHex("#F76707") })
                }),
            widthPoints: 320D,
            heightPoints: 190D));

        Assert.DoesNotContain(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Ellipse &&
            shape.Shape.Width == 5D &&
            shape.Shape.Height == 5D &&
            shape.Shape.FillColor == highlight);
        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Ellipse &&
            shape.Shape.Width == 5D &&
            shape.Shape.Height == 5D &&
            shape.Shape.FillColor == OfficeColor.ParseHex("#F76707"));
    }

    [Fact]
    public void FlowDrawing_UsesSharedChartLayoutForDenseLabelsAndLegend() {
        string[] categories = Enumerable.Range(1, 12).Select(index => "M" + index.ToString("00", System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        OfficeChartSeries[] series = Enumerable.Range(1, 6)
            .Select(index => new OfficeChartSeries(
                "Series " + index.ToString(System.Globalization.CultureInfo.InvariantCulture),
                categories.Select((_, category) => 8D + index * 2D + category).ToArray()))
            .ToArray();
        var layout = new OfficeChartLayout(
            seriesLegendWidthRatio: 0.42D,
            legendRowHeight: 14D,
            legendSwatchSize: 8D,
            legendTextGap: 5D,
            legendFontSize: 8.5D,
            axisLabelFontSize: 7.5D,
            categoryAxisLabelWidth: 44D,
            maximumCategoryAxisLabels: 3);

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Dense chart",
            "Dense Layout",
            OfficeChartKind.Line,
            new OfficeChartData(categories, series),
            widthPoints: 360D,
            heightPoints: 210D,
            layout: layout));

        var textBoxes = drawing.Elements.OfType<OfficeDrawingText>().ToList();
        var categoryLabels = textBoxes.Where(item => categories.Contains(item.Text)).OrderBy(item => item.X).ToList();
        var legendLabels = textBoxes.Where(item => item.Text.StartsWith("Series ", System.StringComparison.Ordinal)).OrderBy(item => item.Y).ToList();

        Assert.Equal(3, categoryLabels.Count);
        Assert.All(categoryLabels, item => Assert.Equal(7.5D, item.Font.Size, 3));
        Assert.Equal(6, legendLabels.Count);
        Assert.All(legendLabels, item => Assert.Equal(8.5D, item.Font.Size, 3));
        OfficeDrawingQualityReport report = OfficeDrawingQualityAnalyzer.Analyze(drawing);
        Assert.False(report.HasIssues, string.Join("; ", report.Issues.Select(issue => issue.ToString())));

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 460,
                PageHeight = 310,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                CompressContentStreams = false
            })
            .Drawing(drawing, PdfAlign.Left)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));

        Assert.Contains("Dense Layout", text, System.StringComparison.Ordinal);
        Assert.Contains("Series 1", text, System.StringComparison.Ordinal);
        Assert.Contains("Series 6", text, System.StringComparison.Ordinal);
        Assert.Contains("M01", text, System.StringComparison.Ordinal);
        Assert.DoesNotContain("M02", text, System.StringComparison.Ordinal);
    }

    [Fact]
    public void FlowDrawing_PreventsDenseCategoryLabelOverlapByDefault() {
        string[] categories = Enumerable.Range(1, 12).Select(index => "M" + index.ToString("00", System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Dense chart",
            "Dense Labels",
            OfficeChartKind.Line,
            new OfficeChartData(
                categories,
                new[] {
                    new OfficeChartSeries("Actual", categories.Select((_, index) => 10D + index).ToArray())
                }),
            widthPoints: 300D,
            heightPoints: 180D,
            layout: new OfficeChartLayout(maximumCategoryAxisLabels: 12)));

        var categoryLabels = drawing.Elements
            .OfType<OfficeDrawingText>()
            .Where(item => categories.Contains(item.Text))
            .ToList();
        OfficeDrawingQualityReport report = OfficeDrawingQualityAnalyzer.Analyze(drawing);

        Assert.True(categoryLabels.Count < categories.Length);
        Assert.False(report.HasIssues, string.Join("; ", report.Issues.Select(issue => issue.ToString())));
    }

    [Fact]
    public void FlowDrawing_StrictDenseCategoryLabelsCanSurfaceQualityIssues() {
        string[] categories = Enumerable.Range(1, 12).Select(index => "M" + index.ToString("00", System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        OfficeChartRenderingResult result = OfficeChartDrawingRenderer.RenderWithQuality(new OfficeChartSnapshot(
            "Strict dense chart",
            "Strict Labels",
            OfficeChartKind.Line,
            new OfficeChartData(
                categories,
                new[] {
                    new OfficeChartSeries("Actual", categories.Select((_, index) => 10D + index).ToArray())
                }),
            widthPoints: 300D,
            heightPoints: 180D,
            layout: new OfficeChartLayout(maximumCategoryAxisLabels: 12, preventLabelOverlap: false)));

        Assert.Contains(result.QualityReport.Issues, issue => issue.Kind == OfficeDrawingQualityIssueKind.TextOverlap);
    }

    [Fact]
    public void FlowDrawing_RendersSharedStyledChartPaletteAndTextColors() {
        var style = new OfficeChartStyle(
            palette: new[] {
                OfficeColor.FromRgb(18, 52, 86),
                OfficeColor.FromRgb(120, 40, 160)
            },
            fontFamily: "Aptos",
            backgroundColor: OfficeColor.FromRgb(242, 248, 255),
            borderColor: OfficeColor.FromRgb(12, 34, 56),
            axisColor: OfficeColor.FromRgb(64, 70, 80),
            gridLineColor: OfficeColor.FromRgb(210, 220, 235),
            textColor: OfficeColor.FromRgb(20, 30, 45),
            mutedTextColor: OfficeColor.FromRgb(90, 105, 125),
            titleColor: OfficeColor.FromRgb(200, 10, 10));
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Styled chart",
            "Styled Revenue",
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeChartSeries("Actual", new[] { 12D, 18D, 24D }),
                    new OfficeChartSeries("Target", new[] { 10D, 20D, 26D })
                }),
            widthPoints: 320D,
            heightPoints: 190D,
            style: style));

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 280,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                CompressContentStreams = false
            })
            .Drawing(drawing, PdfAlign.Left)
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.949 0.973 1 rg", raw, System.StringComparison.Ordinal);
        Assert.Contains("0.071 0.204 0.337 rg", raw, System.StringComparison.Ordinal);
        Assert.Contains("0.471 0.157 0.627 rg", raw, System.StringComparison.Ordinal);
        Assert.Contains("0.784 0.039 0.039 rg", raw, System.StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));

        Assert.Contains("Styled Revenue", text, System.StringComparison.Ordinal);
        Assert.Contains("Actual", text, System.StringComparison.Ordinal);
        Assert.Contains("Target", text, System.StringComparison.Ordinal);
        Assert.Contains("Q1", text, System.StringComparison.Ordinal);
    }

    [Fact]
    public void FlowDrawing_RendersSharedHorizontalBarChartLabels() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Risk bars",
            "Risk Exposure",
            OfficeChartKind.BarClustered,
            new OfficeChartData(
                new[] { "Low", "Medium", "High" },
                new[] {
                    new OfficeChartSeries("Current", new[] { 4D, 9D, 14D }),
                    new OfficeChartSeries("Target", new[] { 3D, 8D, 12D })
                }),
            widthPoints: 320D,
            heightPoints: 190D));

        var categoryLabels = drawing.Elements
            .OfType<OfficeDrawingText>()
            .Where(text => text.Text == "Low" || text.Text == "Medium" || text.Text == "High")
            .ToDictionary(text => text.Text);
        Assert.True(categoryLabels["High"].Y < categoryLabels["Medium"].Y && categoryLabels["Medium"].Y < categoryLabels["Low"].Y, "Expected horizontal bar chart categories to render in Word display order.");

        int verticalGridLines = drawing.Shapes.Count(shape =>
            shape.Shape.Kind == OfficeShapeKind.Line &&
            shape.Shape.StrokeWidth == 0.5D &&
            shape.Shape.Width <= 1D &&
            shape.Shape.Height > 20D);
        int horizontalGridLines = drawing.Shapes.Count(shape =>
            shape.Shape.Kind == OfficeShapeKind.Line &&
            shape.Shape.StrokeWidth == 0.5D &&
            shape.Shape.Width > 20D &&
            shape.Shape.Height <= 1D);
        Assert.Equal(3, verticalGridLines);
        Assert.Equal(0, horizontalGridLines);

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 280,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                CompressContentStreams = false
            })
            .Drawing(drawing, PdfAlign.Left)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));

        Assert.Contains("Risk Exposure", text, System.StringComparison.Ordinal);
        Assert.Contains("Current", text, System.StringComparison.Ordinal);
        Assert.Contains("Target", text, System.StringComparison.Ordinal);
        Assert.Contains("Low", text, System.StringComparison.Ordinal);
        Assert.Contains("Medium", text, System.StringComparison.Ordinal);
        Assert.Contains("High", text, System.StringComparison.Ordinal);
    }

    [Fact]
    public void FlowDrawing_RendersSharedRadarChartCategoryLabels() {
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Capability radar",
            "Capability Radar",
            OfficeChartKind.Radar,
            new OfficeChartData(
                new[] { "Security", "Reliability", "UX", "Speed", "Cost" },
                new[] {
                    new OfficeChartSeries("Current", new[] { 7D, 6D, 5D, 8D, 4D }),
                    new OfficeChartSeries("Target", new[] { 9D, 8D, 7D, 9D, 6D })
                }),
            widthPoints: 340D,
            heightPoints: 210D));

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 440,
                PageHeight = 300,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                CompressContentStreams = false
            })
            .Drawing(drawing, PdfAlign.Left)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));

        Assert.Contains("Capability Radar", text, System.StringComparison.Ordinal);
        Assert.Contains("Current", text, System.StringComparison.Ordinal);
        Assert.Contains("Target", text, System.StringComparison.Ordinal);
        Assert.Contains("Security", text, System.StringComparison.Ordinal);
        Assert.Contains("Reliability", text, System.StringComparison.Ordinal);
        Assert.Contains("Speed", text, System.StringComparison.Ordinal);
    }

    [Fact]
    public void FlowDrawing_RendersRadarPointColorsOnMarkers() {
        OfficeColor highlight = OfficeColor.ParseHex("#F76707");
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Radar point colors",
            "Radar Points",
            OfficeChartKind.Radar,
            new OfficeChartData(
                new[] { "Security", "Reliability", "UX", "Speed" },
                new[] {
                    new OfficeChartSeries(
                        "Current",
                        new[] { 7D, 6D, 5D, 8D },
                        null,
                        OfficeColor.ParseHex("#2563EB"),
                        new OfficeColor?[] { null, null, highlight, null })
                }),
            widthPoints: 300D,
            heightPoints: 190D));

        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Ellipse &&
            shape.Shape.Width == 4D &&
            shape.Shape.Height == 4D &&
            shape.Shape.FillColor == highlight &&
            shape.Shape.StrokeColor == highlight);
    }

}
