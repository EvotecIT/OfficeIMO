using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingAllSeverityBatch20SecurityTests {
    [Fact]
    public void LegacyDrawingEntryPointsRetainExactBinarySignatures() {
        Assert.NotNull(typeof(OfficeImagePatternLayout).GetConstructor(new[] {
            typeof(OfficeImagePlacement),
            typeof(OfficeImagePlacement),
            typeof(bool),
            typeof(bool)
        }));
        Assert.NotNull(typeof(OfficeRichTextLine).GetConstructor(new[] {
            typeof(IReadOnlyList<OfficeRichTextSegment>)
        }));
        Assert.NotNull(typeof(OfficeDrawingText).GetConstructor(new[] {
            typeof(string),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(OfficeFontInfo),
            typeof(OfficeColor?),
            typeof(OfficeTextAlignment),
            typeof(double?)
        }));
        Assert.NotNull(typeof(OfficeDrawing).GetMethod(nameof(OfficeDrawing.AddText), new[] {
            typeof(string),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(double),
            typeof(OfficeFontInfo),
            typeof(OfficeColor?),
            typeof(OfficeTextAlignment),
            typeof(double?)
        }));

        Assert.Contains(typeof(OfficeChartLayout).GetConstructors(), constructor => {
            System.Reflection.ParameterInfo[] parameters = constructor.GetParameters();
            return parameters.Length > 0
                && parameters[0].ParameterType == typeof(double?)
                && parameters[parameters.Length - 1].Name == "verticalAxisMinorTickMark";
        });
        Assert.Contains(typeof(OfficeChartLayout).GetConstructors(), constructor => {
            System.Reflection.ParameterInfo[] parameters = constructor.GetParameters();
            return parameters.Length > 1
                && parameters[0].ParameterType == typeof(bool)
                && parameters[parameters.Length - 1].Name == "verticalAxisMinorTickMark";
        });
    }

    [Fact]
    public void CustomCenterImageFlipsUseTheSharedDestinationTransform() {
        OfficeTransform transform = new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 20D, 10D),
            rotationDegrees: 90D,
            rotationCenterX: 0D,
            rotationCenterY: 0D,
            flipHorizontal: true)
            .CreateUnitSquareTransform();

        Assert.Equal(new OfficeTransform(0D, -20D, -10D, 0D, -20D, -10D), transform);
    }

    [Fact]
    public void SvgCompositionAvoidsCollisionsWithPreexistingNamespacedIds() {
        const string first =
            "<defs><linearGradient id=\"shared\"/><linearGradient id=\"officeimo-layer-1-shared\"/></defs>" +
            "<rect fill=\"url(#shared)\"/>";
        const string second =
            "<defs><linearGradient id=\"shared\"/></defs><rect fill=\"url(#shared)\"/>";

        string svg = OfficeImageComposer.ComposeSvg(
            20,
            10,
            OfficeColor.White,
            new[] {
                OfficeImageLayer.FromSvgInner(first, 0D, 0D, 10D, 10D),
                OfficeImageLayer.FromSvgInner(second, 10D, 0D, 10D, 10D)
            });

        Assert.Contains("id=\"officeimo-layer-1-shared-2\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"url(#officeimo-layer-1-shared-2)\"", svg, StringComparison.Ordinal);
        Assert.Contains("id=\"officeimo-layer-1-shared\"", svg, StringComparison.Ordinal);
        Assert.Contains("id=\"officeimo-layer-2-shared\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void StackedTextCannotUseSingleLineSvgOrRasterFastPaths() {
        var drawing = new OfficeDrawing(40D, 60D);
        drawing.AddText(
            "ABC",
            0D,
            0D,
            40D,
            60D,
            new OfficeFontInfo("Aptos", 12D),
            OfficeColor.Black,
            OfficeTextAlignment.Center,
            stackedText: true);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        int[] paintedRows = Enumerable.Range(0, raster.Height)
            .Where(y => Enumerable.Range(0, raster.Width)
                .Any(x => raster.GetPixel(x, y) != OfficeColor.Transparent))
            .ToArray();

        Assert.Equal(3, CountOccurrences(svg, "<text"));
        Assert.NotEmpty(paintedRows);
        Assert.True(paintedRows[paintedRows.Length - 1] - paintedRows[0] > 20);
    }

    [Fact]
    public void SubTenthPercentSourceCropsRetainTheirExactVisibleSize() {
        var crop = new OfficeImageSourceCrop(0.9995D, 0.99925D, 0D, 0D);

        Assert.True(crop.HasVisibleSourceArea);
        Assert.Equal(0.0005D, crop.VisibleWidth, 12);
        Assert.Equal(0.00075D, crop.VisibleHeight, 12);
    }

    [Fact]
    public void CollinearRadarSeriesStillRendersItsStrokeOutline() {
        OfficeColor seriesColor = OfficeColor.Red;
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(new OfficeChartSnapshot(
            "Degenerate radar",
            null,
            OfficeChartKind.Radar,
            new OfficeChartData(
                new[] { "North", "East", "South", "West" },
                new[] {
                    new OfficeChartSeries(
                        "Series",
                        new[] { 1D, 0D, 1D, 0D },
                        xValues: null,
                        color: seriesColor,
                        pointColors: null,
                        showMarkers: false)
                }),
            widthPoints: 240D,
            heightPoints: 180D,
            layout: new OfficeChartLayout(
                fillRadarSeries: false,
                showLegend: false,
                showMarkers: false)));

        Assert.Contains(drawing.Shapes, shape =>
            shape.Shape.Kind == OfficeShapeKind.Line
            && shape.Shape.StrokeColor == seriesColor);
    }

    [Fact]
    public void UnwrappedLineMeasurementPreservesAuthoredWhitespace() {
        static double Measure(string? value, double size) => (value?.Length ?? 0) * size;

        IReadOnlyList<OfficeTextLine> lines = OfficeTextLayoutEngine.MeasureUnwrappedLines(
            "  A\tB  \r\n C ",
            1D,
            Measure);

        Assert.Equal(new[] { "  A B  ", " C " }, lines.Select(line => line.Text).ToArray());
        Assert.Equal(new[] { 7D, 3D }, lines.Select(line => line.Width).ToArray());
    }

    private static int CountOccurrences(string value, string token) {
        int count = 0;
        int start = 0;
        while ((start = value.IndexOf(token, start, StringComparison.Ordinal)) >= 0) {
            count++;
            start += token.Length;
        }

        return count;
    }
}
