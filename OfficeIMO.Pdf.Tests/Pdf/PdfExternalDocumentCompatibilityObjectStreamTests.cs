using System.IO.Compression;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {

    [Fact]
    public void ReadExternalObjectStream_DoesNotOverwriteExplicitIndirectObjects() {
        byte[] pdf = BuildExternalObjectStreamWithExplicitReplacementPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Explicit object wins", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Packed object stream wins", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_LaterObjectStreamReplacesEarlierCompressedObjects() {
        byte[] pdf = BuildExternalObjectStreamWithCompressedReplacementPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Later object stream wins", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Earlier object stream wins", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_LaterObjectStreamReplacesEarlierExplicitObjects() {
        byte[] pdf = BuildExternalObjectStreamReplacingEarlierExplicitPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Later object stream wins", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Earlier explicit object wins", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_IgnoresMalformedLaterObjectHeaders() {
        byte[] pdf = BuildExternalObjectStreamWithMalformedTrailingPageObjectPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_IgnoresMalformedLaterDictionaryObjects() {
        byte[] pdf = BuildExternalObjectStreamWithMalformedTrailingPageDictionaryPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReadExternalObjectStream_IgnoresParseFailedLaterDictionaryObjects() {
        byte[] pdf = BuildExternalObjectStreamWithParseFailedTrailingPageDictionaryPdf();

        string text = Normalize(PdfTextExtractor.ExtractAllText(pdf));

        Assert.Contains("Object stream page", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RewritePreflight_DetectsCompressedObjectStreamFormMarkers() {
        byte[] pdf = BuildExternalObjectStreamPdf(includeAcroForm: true);
        string rawPdf = Encoding.ASCII.GetString(pdf);

        Assert.DoesNotContain("AcroForm", rawPdf, StringComparison.Ordinal);
        Assert.Contains("Object stream page", Normalize(PdfTextExtractor.ExtractAllText(pdf)), StringComparison.Ordinal);

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);
        Assert.True(preflight.CanRead);
        Assert.False(preflight.CanRewrite);
        Assert.True(preflight.HasRewriteBlocker(PdfRewriteBlockerKind.Forms));

        PdfMutationBlockedException exception = Assert.Throws<PdfMutationBlockedException>(() => PdfPageExtractor.SplitPages(pdf));
        Assert.Contains("FullRewrite.Forms", exception.Plan.BlockerCodes);
    }
}
