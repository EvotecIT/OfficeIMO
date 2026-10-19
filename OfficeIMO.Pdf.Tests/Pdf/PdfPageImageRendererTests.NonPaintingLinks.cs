using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("/Border [0 0 0]", false)]
    [InlineData("/BS << /W 0 >> /Border [0 0 1]", false)]
    [InlineData("/Border [0 0 1]", true)]
    [InlineData("/BS << /W 1 >> /Border [0 0 0]", true)]
    [InlineData("/BS << /S /S >> /Border [0 0 0]", true)]
    [InlineData("", true)]
    public void RenderPage_OnlyNonPaintingLinksMayOmitAnAppearance(string border, bool expectedOmission) {
        byte[] bytes = Encoding.ASCII.GetBytes("%PDF-1.4\n"
            + "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
            + "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj\n"
            + "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Annots [4 0 R] >> endobj\n"
            + "4 0 obj << /Type /Annot /Subtype /Link /Rect [10 10 90 30] " + border
            + " /A << /S /URI /URI (https://example.com/) >> >> endobj\n"
            + "trailer << /Root 1 0 R /Size 5 >>\n%%EOF");
        var page = PdfReadDocument.Open(bytes).Pages[0];
        Assert.Equal(expectedOmission, page.GetRenderCapabilityDiagnostics().Any(
            item => item.Code == PdfRenderCapabilities.AnnotationAppearanceId));
        var options = new PdfImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoLoss = true } };
        if (expectedOmission) {
            var failure = Assert.Throws<OfficeImageExportPolicyException>(() => page.ExportImage(OfficeImageExportFormat.Png, options));
            Assert.Contains(failure.Diagnostics, item => item.Code == PdfRenderCapabilities.AnnotationAppearanceId);
        } else {
            Assert.NotEmpty(page.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
        }
    }
}
