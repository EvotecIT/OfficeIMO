using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Fact]
    public void FluentFillAndFlatten_ReportsConfiguredAppearanceFontOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("OfficeIMO Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] flattened = PdfDocument
            .Open(BuildTextWidgetFormPdf())
            .Forms
            .FillAndFlatten(new Dictionary<string, string> {
                ["Name"] = "office cafe\u0301"
            }, options)
            .ToBytes();

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        Assert.Contains("office cafe", extracted, StringComparison.Ordinal);
        AssertOpenTypeFeatureAppearanceDiagnostics(report);
    }

    [Fact]
    public void TryFillAndFlatten_ReportsConfiguredAppearanceFontOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFont("OfficeIMO Fill CFF Font", File.ReadAllBytes(fontPath!))
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        PdfOperationResult<PdfDocument> result = PdfDocument
            .Open(BuildTextWidgetFormPdf())
            .Forms
            .TryFillAndFlatten(new Dictionary<string, string> {
                ["Name"] = "office cafe\u0301"
            }, options, null);

        Assert.True(result.CanAttempt);
        Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics));

        byte[] flattened = result.RequireValue().ToBytes();
        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        Assert.Contains("office cafe", extracted, StringComparison.Ordinal);
        AssertOpenTypeFeatureAppearanceDiagnostics(report);
    }

    [Fact]
    public void FluentFillAndFlatten_ReportsFallbackAppearanceFontOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        var fallbackSet = new PdfEmbeddedFontFallbackSet(
            new[] {
                new PdfEmbeddedFontFallbackCandidate("Fallback Fill CFF Font", System.IO.File.ReadAllBytes(fontPath!))
            },
            new[] { PdfStandardFont.Helvetica });
        var report = new PdfConversionReport();
        var options = new PdfFormFillerOptions()
            .UseAppearanceFontFallbacks(fallbackSet)
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

        byte[] flattened = PdfDocument
            .Open(BuildTextWidgetFormPdf())
            .Forms
            .FillAndFlatten(new Dictionary<string, string> {
                ["Name"] = "office cafe\u0301"
            }, options)
            .ToBytes();

        string output = Encoding.ASCII.GetString(flattened);
        string extracted = PdfReadDocument.Open(flattened).ExtractText();

        Assert.DoesNotContain("/AcroForm", output, StringComparison.Ordinal);
        Assert.Contains("/FontFile3", output, StringComparison.Ordinal);
        Assert.Contains("office cafe", extracted, StringComparison.Ordinal);
        AssertOpenTypeFeatureAppearanceDiagnostics(report);
    }

    [Fact]
    public void PathInputStreamOutputFillAndFlatten_ReportsConfiguredAppearanceFontOpenTypeFeatureDiagnostics() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        Assert.NotNull(fontPath);

        string inputPath = Path.Combine(Path.GetTempPath(), "officeimo-form-cff-fill-flatten-input-" + Guid.NewGuid().ToString("N") + ".pdf");
        try {
            File.WriteAllBytes(inputPath, BuildTextWidgetFormPdf());
            var report = new PdfConversionReport();
            var options = new PdfFormFillerOptions()
                .UseAppearanceFont("OfficeIMO Fill CFF Font", File.ReadAllBytes(fontPath!))
                .ReportDiagnosticsTo(report, "OfficeIMO.Tests");

            using var output = new MemoryStream();
            output.WriteByte(31);
            PdfFormFiller.FillAndFlattenFields(inputPath, output, new Dictionary<string, string> {
                ["Name"] = "office cafe\u0301"
            }, options);

            byte[] flattened = SliceAfterPrefix(output, 1);
            string pdfText = Encoding.ASCII.GetString(flattened);
            string extracted = PdfReadDocument.Open(flattened).ExtractText();

            Assert.DoesNotContain("/AcroForm", pdfText, StringComparison.Ordinal);
            Assert.Contains("/FontFile3", pdfText, StringComparison.Ordinal);
            Assert.Contains("office cafe", extracted, StringComparison.Ordinal);
            AssertOpenTypeFeatureAppearanceDiagnostics(report);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
        }
    }
}
