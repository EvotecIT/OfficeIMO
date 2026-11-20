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
        using WordDocument document = CreateChart(errorBars: false);

        OfficeIMO.Pdf.PdfConversionReport report = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false }).Report;

        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "native-body-chart-unsupported");
    }

    [Fact]
    public void StrictPdfNamesWhyAChartIsSkippedAndBestEffortDrawsIt() {
        using WordDocument document = CreateChart(errorBars: true);

        OfficeIMO.Pdf.PdfConversionReport strict = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false }).Report;
        OfficeIMO.Pdf.PdfConversionReport bestEffort = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort }
        }).Report;

        Assert.Contains(strict.Warnings, warning => warning.Code == "native-body-chart-unsupported" &&
            warning.Message.IndexOf("Error bars cannot be projected.", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(bestEffort.Warnings, warning => warning.Code == "native-body-chart-unsupported");
        Assert.Contains(bestEffort.Warnings, warning => warning.Code == OfficeDiagnosticCodes.ChartApproximation &&
            warning.Message == "Error bars are left out.");
    }

    [Fact]
    public void PageImageExportHonoursBestEffortChartFidelity() {
        using WordDocument document = CreateChart(errorBars: true);

        OfficeImageExportResult strict = document.ToImage(new WordImageExportOptions());
        OfficeImageExportResult bestEffort = document.ToImage(new WordImageExportOptions { ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort } });

        Assert.Contains(strict.Diagnostics, diagnostic => diagnostic.Code == "native-body-chart-unsupported");
        Assert.DoesNotContain(bestEffort.Diagnostics, diagnostic => diagnostic.Code == "native-body-chart-unsupported");
        Assert.Contains(bestEffort.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation);
    }

    // Excel and Word write this chart-area fill and outline on every new chart.
    private static WordDocument CreateChart(bool errorBars) {
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
        space.InsertAfter(new C.ShapeProperties(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Background1 }), outline, new A.EffectList()),
            space.GetFirstChild<C.Chart>());
        if (errorBars) {
            C.BarChartSeries series = space.Descendants<C.BarChartSeries>().Single();
            series.InsertBefore(new C.ErrorBars(new C.ErrorBarType { Val = C.ErrorBarValues.Both },
                new C.ErrorBarValueType { Val = C.ErrorValues.StandardError }, new C.NoEndCap { Val = false }),
                series.GetFirstChild<C.CategoryAxisData>());
        }
        return document;
    }
}
