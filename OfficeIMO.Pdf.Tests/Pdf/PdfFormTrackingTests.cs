using System.Collections.Generic;
using System.Text;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFormFillerTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FillFields_PaintsTrackedEmbeddedAndFallbackFontsWithMeasuredAdvances(bool fallback) {
        byte[] data = ManagedTextShapingTestAssets.CreateTrackingFont(PdfFontTrackingTests.Table());
        var options = new PdfFormFillerOptions();
        if (fallback) options.UseAppearanceFontFallbacks(new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Tracked", data) },
            new[] { PdfStandardFont.Helvetica }));
        else options.UseAppearanceFont("Tracked", data);
        byte[] filled = PdfFormFiller.FillFields(
            BuildTextWidgetFormPdfWithDefaultAppearance("/Helv 15 Tf 0 g"),
            new Dictionary<string, string> { ["Name"] = "AB" }, options);
        Assert.Equal("AB", Assert.Single(PdfInspector.Inspect(filled).FormFields).Value);
        Assert.Contains("<0001> 50] TJ", Encoding.ASCII.GetString(filled));
    }
}
