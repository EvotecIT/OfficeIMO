using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIndirectObjectHeaderScannerTests {
    [Fact]
    public void SkipsOverflowedHeaderCandidateAndFindsFollowingObject() {
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n" +
            new string('9', 128 * 1024) +
            " 0 obj\nnull\nendobj\n" +
            "7 0 obj\n42\nendobj\n%%EOF\n");

        var (objects, _) = PdfSyntax.ParseObjects(pdf);

        PdfIndirectObject parsed = Assert.Single(objects).Value;
        Assert.Equal(7, parsed.ObjectNumber);
        Assert.Equal(42D, Assert.IsType<PdfNumber>(parsed.Value).Value);
    }

    [Fact]
    public void AcceptsLargeWhitespaceRunsAtHeaderTokenBoundaries() {
        string whitespace = new string(' ', 64 * 1024);
        byte[] pdf = Encoding.ASCII.GetBytes(
            "%PDF-1.7\n12" +
            whitespace +
            "0" +
            whitespace +
            "obj\n13\nendobj\n%%EOF\n");

        var (objects, _) = PdfSyntax.ParseObjects(
            pdf,
            new PdfReadOptions {
                Limits = new PdfReadLimits {
                    MaxInputBytes = pdf.Length,
                    MaxObjectParsingTime = TimeSpan.FromSeconds(2)
                }
            });

        PdfIndirectObject parsed = Assert.Single(objects).Value;
        Assert.Equal(12, parsed.ObjectNumber);
        Assert.Equal(13D, Assert.IsType<PdfNumber>(parsed.Value).Value);
    }
}
