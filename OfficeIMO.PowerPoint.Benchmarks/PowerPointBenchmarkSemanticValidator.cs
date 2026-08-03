using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Validates the shared benchmark corpus by content and formatting so both
/// producer lanes measure equivalent work.
/// </summary>
internal static class PowerPointBenchmarkSemanticValidator {
    internal static void Validate(PresentationDocument document,
        int expectedSlideCount, string operation) {
        PresentationPart presentationPart = document.PresentationPart
            ?? throw new InvalidOperationException(
                operation + " produced no presentation part.");
        P.Presentation presentation = presentationPart.Presentation
            ?? throw new InvalidOperationException(
                operation + " produced no presentation root.");
        P.SlideId[] slideIds = presentation.SlideIdList?
            .Elements<P.SlideId>().ToArray() ?? Array.Empty<P.SlideId>();
        if (slideIds.Length != expectedSlideCount) {
            throw new InvalidOperationException(
                $"{operation} produced {slideIds.Length} slides; expected {expectedSlideCount}.");
        }

        for (int index = 0; index < slideIds.Length; index++) {
            string relationshipId = slideIds[index].RelationshipId?.Value
                ?? throw new InvalidOperationException(
                    $"{operation} slide {index + 1} has no relationship id.");
            SlidePart slidePart = (SlidePart)presentationPart.GetPartById(
                relationshipId);
            ValidateSlide(slidePart, index, expectedSlideCount, operation);
        }
    }

    private static void ValidateSlide(SlidePart slidePart, int index,
        int slideCount, string operation) {
        P.Slide slide = slidePart.Slide
            ?? throw new InvalidOperationException(
                $"{operation} slide {index + 1} has no slide root.");
        string[] texts = slide.Descendants<A.Text>()
            .Select(text => text.Text ?? string.Empty).ToArray();

        RequireText(texts, $"Operational review {index + 1}", index, operation);
        RequireText(texts,
            $"Slide {index + 1} of {slideCount} · deterministic benchmark corpus",
            index, operation);
        RequireText(texts, "OfficeIMO.PowerPoint performance corpus", index,
            operation);
        ValidateTitleStyle(slide, index, operation);
        ValidateBackground(slide, index, operation);
        ValidateMetricCards(slide, index, operation);

        if (index % 3 == 0) {
            ValidateTable(slide, index, operation);
        } else {
            ValidateDetailPanels(texts, index, operation);
        }

        if (index % 5 == 0) {
            ValidateChart(slidePart, index, operation);
        } else {
            RequireText(texts,
                "Measured work includes editable text, vector shapes, package serialization, and rendering.",
                index, operation);
        }

        int markerCount = texts.Count(text => string.Equals(text, "Reviewed",
            StringComparison.Ordinal));
        int expectedMarkers = string.Equals(operation, "OpenEditSave",
            StringComparison.OrdinalIgnoreCase) && index % 10 == 0 ? 1 : 0;
        if (markerCount != expectedMarkers) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} contains {markerCount} review marker(s); expected {expectedMarkers}.");
        }
    }

    private static void ValidateTitleStyle(P.Slide slide, int index,
        string operation) {
        P.Shape? title = FindShapeByText(slide,
            $"Operational review {index + 1}");
        if (title == null) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} lost its title shape.");
        }

        bool hasExpectedRun = title.Descendants<A.RunProperties>().Any(run =>
                run.FontSize?.Value == 2400 && run.Bold?.Value == true
                && HasRgbFill(run, "0F172A"))
            || title.Descendants<A.DefaultRunProperties>().Any(run =>
                run.FontSize?.Value == 2400 && run.Bold?.Value == true
                && HasRgbFill(run, "0F172A"));
        if (!hasExpectedRun) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} lost the title's 24pt bold #0F172A styling.");
        }
    }

    private static void ValidateBackground(P.Slide slide, int index,
        string operation) {
        string expected = index % 2 == 0 ? "F8FAFC" : "F1F5F9";
        string? actual = slide.CommonSlideData?.Background?
            .BackgroundProperties?.GetFirstChild<A.SolidFill>()?
            .RgbColorModelHex?.Val?.Value;
        if (!string.Equals(actual, expected,
                StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} background is '{actual ?? "missing"}'; expected {expected}.");
        }
    }

    private static void ValidateMetricCards(P.Slide slide, int index,
        string operation) {
        string[] colors = { "DBEAFE", "DCFCE7", "FEF3C7", "FCE7F3" };
        for (int card = 0; card < 4; card++) {
            P.Shape? shape = FindShapeAt(slide, 40D + card * 220D, 120D);
            string expectedColor = colors[(card + index) % colors.Length];
            string? actualColor = shape?.ShapeProperties?
                .GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value;
            if (!string.Equals(actualColor, expectedColor,
                    StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidOperationException(
                    $"{operation} slide {index + 1} lost metric card {card + 1} fill {expectedColor}.");
            }
        }
    }

    private static void ValidateTable(P.Slide slide, int index,
        string operation) {
        A.Table table = slide.Descendants<A.Table>().SingleOrDefault()
            ?? throw new InvalidOperationException(
                $"{operation} slide {index + 1} lost its benchmark table.");
        P.GraphicFrame frame = table.Ancestors<P.GraphicFrame>().Single();
        A.Offset? offset = frame.Transform?.Offset;
        A.Extents? extents = frame.Transform?.Extents;
        const long emusPerPoint = 12700L;
        if (offset?.X?.Value != 40L * emusPerPoint
            || offset.Y?.Value != 224L * emusPerPoint
            || extents?.Cx?.Value != 300L * emusPerPoint
            || extents.Cy?.Value != 220L * emusPerPoint) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} table is not positioned at 40,224 with a 300x220 point extent.");
        }
        string[,] expected = {
            { "Metric", "Current", "Target" },
            { "Quality", (92 + index % 7).ToString(CultureInfo.InvariantCulture), "98" },
            { "Coverage", (80 + index % 15).ToString(CultureInfo.InvariantCulture), "95" },
            { "Latency", (24 + index % 9).ToString(CultureInfo.InvariantCulture), "20" }
        };
        A.TableRow[] rows = table.Elements<A.TableRow>().ToArray();
        if (rows.Length != 4 || rows.Any(row =>
                row.Elements<A.TableCell>().Count() != 3)) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} table is not 4x3.");
        }
        for (int row = 0; row < 4; row++) {
            A.TableCell[] cells = rows[row].Elements<A.TableCell>().ToArray();
            for (int column = 0; column < 3; column++) {
                string actual = string.Concat(cells[column]
                    .Descendants<A.Text>().Select(text => text.Text));
                if (!string.Equals(actual, expected[row, column],
                        StringComparison.Ordinal)) {
                    throw new InvalidOperationException(
                        $"{operation} slide {index + 1} table cell [{row},{column}] is '{actual}'; expected '{expected[row, column]}'.");
                }
            }
        }
        foreach (A.TableCell header in rows[0].Elements<A.TableCell>()) {
            string? fill = header.TableCellProperties?
                .GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value;
            bool bold = header.Descendants<A.RunProperties>()
                    .Any(run => run.Bold?.Value == true)
                || header.Descendants<A.DefaultRunProperties>()
                    .Any(run => run.Bold?.Value == true);
            if (!string.Equals(fill, "DBEAFE",
                    StringComparison.OrdinalIgnoreCase) || !bold) {
                throw new InvalidOperationException(
                    $"{operation} slide {index + 1} lost its styled table header.");
            }
        }
    }

    private static void ValidateDetailPanels(IReadOnlyCollection<string> texts,
        int index, string operation) {
        for (int row = 0; row < 3; row++) {
            RequireText(texts,
                $"Workstream {row + 1}: checkpoint {index + row + 1}",
                index, operation);
        }
    }

    private static void ValidateChart(SlidePart slidePart, int index,
        string operation) {
        ChartPart chartPart = slidePart.ChartParts.SingleOrDefault()
            ?? throw new InvalidOperationException(
                $"{operation} slide {index + 1} lost its chart.");
        C.Chart chart = chartPart.ChartSpace?.GetFirstChild<C.Chart>()
            ?? throw new InvalidOperationException(
                $"{operation} slide {index + 1} chart has no root.");
        if (chart.GetFirstChild<C.Title>() != null) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} chart unexpectedly has a title.");
        }
        C.BarChart barChart = chart.GetFirstChild<C.PlotArea>()?
            .Elements<C.BarChart>().SingleOrDefault()
            ?? throw new InvalidOperationException(
                $"{operation} slide {index + 1} chart is not a bar chart.");
        if (barChart.BarDirection?.Val?.Value != C.BarDirectionValues.Bar
            || barChart.BarGrouping?.Val?.Value
                != C.BarGroupingValues.Clustered) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} chart is not a horizontal clustered bar chart.");
        }
        C.BarChartSeries[] series = barChart.Elements<C.BarChartSeries>()
            .ToArray();
        if (series.Length != 2) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} chart has {series.Length} series; expected 2.");
        }
        string[] expectedNames = { "Actual", "Target" };
        double[][] expectedValues = {
            new[] { 12D + index, 18D + index, 24D + index, 30D + index },
            new[] { 15D + index, 20D + index, 26D + index, 32D + index }
        };
        for (int seriesIndex = 0; seriesIndex < series.Length; seriesIndex++) {
            string name = FirstCachedValue(series[seriesIndex].SeriesText);
            if (!string.Equals(name, expectedNames[seriesIndex],
                    StringComparison.Ordinal)) {
                throw new InvalidOperationException(
                    $"{operation} slide {index + 1} chart series {seriesIndex} is '{name}'; expected '{expectedNames[seriesIndex]}'.");
            }
            string[] categories = CachedValues(
                series[seriesIndex].GetFirstChild<C.CategoryAxisData>());
            if (!categories.SequenceEqual(new[] { "Q1", "Q2", "Q3", "Q4" },
                    StringComparer.Ordinal)) {
                throw new InvalidOperationException(
                    $"{operation} slide {index + 1} chart lost its categories.");
            }
            double[] values = CachedValues(
                    series[seriesIndex].GetFirstChild<C.Values>())
                .Select(value => double.Parse(value,
                    CultureInfo.InvariantCulture)).ToArray();
            if (!values.SequenceEqual(expectedValues[seriesIndex])) {
                throw new InvalidOperationException(
                    $"{operation} slide {index + 1} chart series '{name}' has unexpected values.");
            }
        }
    }

    private static string FirstCachedValue(OpenXmlElement? element) =>
        CachedValues(element).FirstOrDefault() ?? string.Empty;

    private static string[] CachedValues(OpenXmlElement? element) =>
        element?.Descendants<C.NumericValue>()
            .Select(value => value.Text ?? string.Empty).ToArray()
        ?? Array.Empty<string>();

    private static P.Shape? FindShapeByText(P.Slide slide, string expected) =>
        slide.Descendants<P.Shape>().FirstOrDefault(shape =>
            string.Equals(string.Concat(shape.Descendants<A.Text>()
                .Select(text => text.Text)), expected, StringComparison.Ordinal));

    private static P.Shape? FindShapeAt(P.Slide slide, double leftPoints,
        double topPoints) {
        const double emusPerPoint = 12700D;
        long expectedX = checked((long)Math.Round(leftPoints * emusPerPoint,
            MidpointRounding.AwayFromZero));
        long expectedY = checked((long)Math.Round(topPoints * emusPerPoint,
            MidpointRounding.AwayFromZero));
        return slide.Descendants<P.Shape>().FirstOrDefault(shape => {
            A.Offset? offset = shape.ShapeProperties?
                .GetFirstChild<A.Transform2D>()?.Offset;
            return offset?.X?.Value == expectedX && offset.Y?.Value == expectedY;
        });
    }

    private static bool HasRgbFill(OpenXmlCompositeElement properties,
        string expected) => string.Equals(properties
            .GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value,
            expected, StringComparison.OrdinalIgnoreCase);

    private static void RequireText(IEnumerable<string> texts, string expected,
        int index, string operation) {
        if (!texts.Contains(expected, StringComparer.Ordinal)) {
            throw new InvalidOperationException(
                $"{operation} slide {index + 1} lost expected text '{expected}'.");
        }
    }
}
