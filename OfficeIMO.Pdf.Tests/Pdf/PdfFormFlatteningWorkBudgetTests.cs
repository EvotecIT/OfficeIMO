using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfFormFlatteningWorkBudgetTests {
    [Theory]
    [InlineData(700, 2_000, PdfReadLimitKind.PageContentBytes)]
    [InlineData(2_000, 700, PdfReadLimitKind.RetainedContentBytes)]
    public void SharedAppearanceInvocationsChargeInterpretationWork(
        int pageLimit, long totalLimit, PdfReadLimitKind expectedLimit) {
        byte[] pdf = BuildFormWithRepeatedAppearanceChild();
        var options = new PdfLoadOptions {
            Limits = new PdfReadLimits {
                MaxPageContentBytes = pageLimit,
                MaxRetainedContentBytes = totalLimit
            }
        };

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() =>
            PdfDocument.Load(pdf, options).Forms.Flatten());
        Assert.Equal(expectedLimit, error.Kind);
    }

    [Fact]
    public void SharedAppearanceWithinWorkBudgetStillFlattens() {
        byte[] pdf = BuildFormWithRepeatedAppearanceChild();
        PdfDocument flattened = PdfDocument.Load(pdf).Forms.Flatten();

        Assert.False(PdfInspector.Inspect(flattened.ToBytes()).HasForms);
    }

    private static byte[] BuildFormWithRepeatedAppearanceChild() {
        string parentContent = string.Concat(Enumerable.Repeat("/Shared Do ", 4));
        string childContent = new(' ', 256);
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj", "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Annots [6 0 R] >>", "endobj",
            "5 0 obj", "<< /Fields [6 0 R] >>", "endobj",
            "6 0 obj", "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Field) /Rect [20 20 100 40] /P 3 0 R /AP << /N 9 0 R >> >>", "endobj",
            "9 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 80 20] /Resources << /XObject << /Shared 10 0 R >> >> /Length " + parentContent.Length + " >>",
            "stream", parentContent, "endstream", "endobj",
            "10 0 obj", "<< /Type /XObject /Subtype /Form /BBox [0 0 80 20] /Length " + childContent.Length + " >>",
            "stream", childContent, "endstream", "endobj",
            "trailer", "<< /Root 1 0 R /Size 11 >>", "%%EOF"
        }));
    }
}
