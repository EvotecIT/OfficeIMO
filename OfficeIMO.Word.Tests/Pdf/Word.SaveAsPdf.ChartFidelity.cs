using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

/// <summary>Word PDF and page-image export of charts written with Office's default formatting.</summary>
public sealed class WordChartFidelityPdfTests {
    [Fact]
    public void OfficeChartAreaOutlineIsDrawnInPdf() {
        using WordDocument document = CreateChart(gradientFill: false);

        OfficeIMO.Pdf.PdfConversionReport report = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false }).Report;

        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "native-body-chart-unsupported");
    }

    [Fact]
    public void StrictPdfNamesWhyAChartIsSkippedAndBestEffortDrawsIt() {
        using WordDocument document = CreateChart(gradientFill: true);

        OfficeIMO.Pdf.PdfConversionReport strict = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false }).Report;
        OfficeIMO.Pdf.PdfConversionReport bestEffort = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort }
        }).Report;

        Assert.Contains(strict.Warnings, warning => warning.Code == "native-body-chart-unsupported" &&
            warning.Message.IndexOf("series fill, line or effect", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(bestEffort.Warnings, warning => warning.Code == "native-body-chart-unsupported");
        Assert.Contains(bestEffort.Warnings, warning => warning.Code == OfficeDiagnosticCodes.ChartApproximation &&
            warning.Message.EndsWith("it is drawn in its plain series colour.", StringComparison.Ordinal));
    }

    [Fact]
    public void PageImageExportHonoursBestEffortChartFidelity() {
        using WordDocument document = CreateChart(gradientFill: true);

        OfficeImageExportResult strict = document.ToImage(new WordImageExportOptions());
        OfficeImageExportResult bestEffort = document.ToImage(new WordImageExportOptions { ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort } });

        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "native-body-chart-unsupported");
        Assert.DoesNotContain(bestEffort.Diagnostics, diagnostic => diagnostic.Code == "native-body-chart-unsupported");
        Assert.Contains(bestEffort.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation);
    }

    [Fact]
    public void ErrorBarsOnAPlainChartAreDrawnInStrictPdf() {
        using WordDocument document = CreateChart(gradientFill: false, errorBars: true, chartArea: false);

        OfficeIMO.Pdf.PdfConversionReport strict = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false }).Report;

        Assert.DoesNotContain(strict.Warnings, warning => warning.Code == "native-body-chart-unsupported");
        Assert.DoesNotContain(strict.Warnings, warning => warning.Code == OfficeDiagnosticCodes.ChartApproximation);
    }

    [Fact]
    public void BestEffortDrawsAManuallyPlacedSecondaryAxisTitleInItsAutomaticPosition() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeSeries("Count", new[] { 100d }),
            new OfficeSeries("Ratio", new[] { 1d }, null, null, null, true, renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }));
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout().WithTitle("Ratio"));
        chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right)
            .GetFirstChild<C.Title>()!.GetFirstChild<C.Layout>()!.Append(new C.ManualLayout(new C.Left { Val = .3 }));

        AssertStrictSkipsAndBestEffortDraws(document);
    }

    [Fact]
    public void BestEffortDrawsAxesWhoseLabelColoursDiffer() {
        using WordDocument document = CreateChart(gradientFill: false);
        C.PlotArea plot = document.Charts[0].ChartPart!.ChartSpace!.Descendants<C.PlotArea>().Single();
        plot.GetFirstChild<C.CategoryAxis>()!.AddChild(AxisText("C00000"), true);
        plot.GetFirstChild<C.ValueAxis>()!.AddChild(AxisText("0070C0"), true);

        AssertStrictSkipsAndBestEffortDraws(document);
    }

    [Fact]
    public void BestEffortDrawsAPointWithMixedCustomTextAsAStandardLabel() {
        using WordDocument document = CreateChart(gradientFill: false);
        C.DataLabels labels = document.Charts[0].ChartPart!.ChartSpace!.Descendants<C.DataLabels>().First();
        labels.GetFirstChild<C.ShowValue>()!.Val = true;
        labels.PrependChild(new C.DataLabel(new C.Index { Val = 0U }, new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.Run(new A.RunProperties { Bold = true }, new A.Text("Bold")),
                new A.Run(new A.RunProperties { Bold = false }, new A.Text(" plain"))))), new C.ShowValue { Val = true }));

        OfficeIMO.Pdf.PdfConversionReport bestEffort = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort }
        }).Report;

        Assert.DoesNotContain(bestEffort.Warnings, warning => warning.Code == "native-body-chart-unsupported");
        Assert.Contains(bestEffort.Warnings, warning => warning.Code == OfficeDiagnosticCodes.ChartApproximation &&
            warning.Message == "A data label with its own text or position is drawn as a standard label.");
    }

    [Fact]
    public void ErrorBarsOnADataTableChartAreReportedAndBestEffortDrawsTheChart() {
        using WordDocument document = CreateChart(gradientFill: false, errorBars: true);
        document.Charts[0].ChartPart!.ChartSpace!.Descendants<C.PlotArea>().Single()
            .AddChild(new C.DataTable(new C.ShowHorizontalBorder { Val = true }), true);

        AssertStrictSkipsAndBestEffortDraws(document);
    }

    [Fact]
    public void BestEffortDrawsATitleThatMixesFonts() {
        using WordDocument document = CreateChart(gradientFill: false);
        C.Chart chart = document.Charts[0].ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        chart.GetFirstChild<C.Title>()?.Remove();
        chart.GetFirstChild<C.AutoTitleDeleted>()?.Remove();
        chart.PrependChild(new C.Title(new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.Run(new A.RunProperties(new A.LatinFont { Typeface = "Arial" }), new A.Text("Sales ")),
                new A.Run(new A.RunProperties(new A.LatinFont { Typeface = "Calibri" }), new A.Text("2026"))))),
            new C.Overlay { Val = false }));

        AssertStrictSkipsAndBestEffortDraws(document);
    }

    [Fact]
    public void LinkedCategoryAxisUsesTheCategoryCacheFormat() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartType.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Revenue", new[] { 4.3d, 2.5d }) }));
        C.ChartSpace space = chart.ChartPart!.ChartSpace!;
        Assert.NotNull(space.GetFirstChild<C.ExternalData>());
        C.CategoryAxisData categories = space.Descendants<C.CategoryAxisData>().Single();
        categories.RemoveAllChildren();
        categories.Append(new C.NumberReference(new C.Formula("Sheet1!$A$2:$A$3"), new C.NumberingCache(new C.FormatCode("yyyy-mm-dd"),
            new C.PointCount { Val = 2U }, new C.NumericPoint(new C.NumericValue("45292")) { Index = 0U },
            new C.NumericPoint(new C.NumericValue("45293")) { Index = 1U })));
        C.CategoryAxis axis = space.Descendants<C.CategoryAxis>().Single();
        axis.RemoveAllChildren<C.NumberingFormat>();
        axis.AddChild(new C.NumberingFormat { FormatCode = "General", SourceLinked = true }, true);

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("yyyy-mm-dd", snapshot.Layout!.CategoryAxisNumberFormat);
    }

    private static void AssertStrictSkipsAndBestEffortDraws(WordDocument document) {
        OfficeIMO.Pdf.PdfConversionReport strict = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false }).Report;
        OfficeIMO.Pdf.PdfConversionReport bestEffort = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort }
        }).Report;

        Assert.Contains(strict.Warnings, warning => warning.Code == "native-body-chart-unsupported");
        Assert.DoesNotContain(bestEffort.Warnings, warning => warning.Code == "native-body-chart-unsupported");
        Assert.Contains(bestEffort.Warnings, warning => warning.Code == OfficeDiagnosticCodes.ChartApproximation);
    }

    private static C.TextProperties AxisText(string color) => new(new A.BodyProperties(), new A.ListStyle(),
        new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = color })))));

    // Excel and Word write this chart-area fill and outline on every new chart.
    private static WordDocument CreateChart(bool gradientFill, bool errorBars = false, bool chartArea = true) {
        WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartType.ColumnClustered,
            new OfficeChartData(new[] { "North", "South" }, new[] { new OfficeSeries("Revenue", new[] { 4.3d, 2.5d }) }));
        C.ChartSpace space = chart.ChartPart!.ChartSpace!;
        space.GetFirstChild<C.ShapeProperties>()?.Remove();
        var outline = new A.Outline(
            new A.SolidFill(new A.SchemeColor(new A.LuminanceModulation { Val = 15000 }, new A.LuminanceOffset { Val = 85000 }) {
                Val = A.SchemeColorValues.Text1
            }),
            new A.Round()) {
            Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center
        };
        if (chartArea)
            space.InsertAfter(new C.ShapeProperties(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Background1 }), outline, new A.EffectList()),
                space.GetFirstChild<C.Chart>());
        if (errorBars) {
            C.BarChartSeries series = space.Descendants<C.BarChartSeries>().Single();
            series.InsertBefore(new C.ErrorBars(new C.ErrorBarType { Val = C.ErrorBarValues.Both },
                new C.ErrorBarValueType { Val = C.ErrorValues.StandardError }, new C.NoEndCap { Val = false }),
                series.GetFirstChild<C.CategoryAxisData>());
        }
        if (gradientFill) {
            C.BarChartSeries series = space.Descendants<C.BarChartSeries>().Single();
            series.GetFirstChild<C.ChartShapeProperties>()?.Remove();
            series.InsertAfter(new C.ChartShapeProperties(new A.GradientFill(new A.GradientStopList(
                new A.GradientStop(new A.RgbColorModelHex { Val = "112233" }) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = "445566" }) { Position = 100000 }))),
                series.GetFirstChild<C.SeriesText>());
        }
        return document;
    }
}
