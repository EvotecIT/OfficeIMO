using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void RedactingAMarkedNoShowObjectPreservesStateUsedByUnrelatedText(bool metadata, bool precise) {
        string marker = metadata ? "/Span << /ActualText (metadata-value) >> BDC" : "/Span BMC";
        byte[] bytes = MarkedContentSecurityPdf("BT " + marker + " /F1 24 Tf 2 Tc EMC ET " +
            "BT 20 170 Td (Public summary) Tj ET BT /F1 12 Tf 20 20 Td (secret) Tj ET");
        var before = PdfReadDocument.Open(bytes).Pages[0];
        var original = Assert.Single(before.GetTextSpans(), span => span.Text == "Public summary");
        byte[] result;
        if (precise) {
            var source = PdfDocument.Load(bytes);
            var plan = source.Redactions.Search(new PdfRedactionSearchOptions {
                TextSelection = PdfRedactionTextSelection.MatchedGlyphs
            }.AddLiteral("secret"));
            result = source.Redactions.Apply(plan).ToBytes();
        } else {
            var secret = Assert.Single(before.GetTextSpans(), span => span.Text == "secret");
            var bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(secret);
            result = PdfRedactionApplier.Apply(bytes, new[] {
                new PdfRedactionArea(1, bounds.Left - 1, bounds.Bottom - 1, bounds.Width + 2, bounds.Height + 2)
            });
        }

        var retained = Assert.Single(PdfReadDocument.Open(result).Pages[0].GetTextSpans(), span => span.Text == "Public summary");
        Assert.Equal(original.FontSize, retained.FontSize);
        Assert.Equal(original.Advance, retained.Advance, 6);
        Assert.DoesNotContain("secret", PdfTextExtractor.ExtractAllText(result), System.StringComparison.Ordinal);
        if (metadata) Assert.DoesNotContain("metadata-value", PdfEncoding.Latin1GetString(result), System.StringComparison.Ordinal);
    }

    [Fact]
    public void RedactionReparsesHonorARaisedMarkedContentDepthLimit() {
        byte[] bytes = MarkedContentSecurityPdf("BT /F1 12 Tf " +
            string.Concat(Enumerable.Repeat("/Span BMC ", 129)) + "20 100 Td (secret) Tj " +
            string.Concat(Enumerable.Repeat("EMC ", 129)) + "ET");
        var options = new PdfLoadOptions { Limits = new PdfReadLimits { MaxContentNestingDepth = 256 } };
        Assert.Contains("secret", PdfReadDocument.Open(bytes, options).ExtractText());

        byte[] result = PdfRedactionApplier.Apply(bytes, new[] { new PdfRedactionArea(1, 0, 0, 200, 200) }, readOptions: options);

        Assert.DoesNotContain("secret", PdfReadDocument.Open(result, options).ExtractText(), System.StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(2, true)]
    [InlineData(3, true)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    public void ActualTextDrawingScopesRespectTheContentNestingLimit(int depth, bool actualText) {
        string content = string.Concat(Enumerable.Repeat(actualText ? "/Span << /ActualText (x) >> BDC " : "/Span BMC ", depth)) +
            "1 0 0 rg 0 0 20 20 re f " + string.Concat(Enumerable.Repeat("EMC ", depth));
        var page = PdfReadDocument.Open(MarkedContentSecurityPdf(content), new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 2 }
        }).Pages[0];

        if (depth > 2) {
            var error = Assert.Throws<PdfReadLimitException>(() => page.ToDrawing());
            Assert.Equal(PdfReadLimitKind.ContentNestingDepth, error.Kind);
            Assert.Equal(2, error.Limit);
            Assert.Equal(3, error.Actual);
        } else {
            Assert.Equal(OfficeColor.Red, OfficeDrawingRasterRenderer.Render(page.ToDrawing()).GetPixel(10, 190));
        }
    }

    [Fact]
    public void DisjointMarkedContentScopesReleaseTheirNestingDepth() {
        string content = string.Concat(Enumerable.Repeat("/Span BMC /Span BMC 1 0 0 rg 0 0 20 20 re f EMC EMC ", 5));
        var page = PdfReadDocument.Open(MarkedContentSecurityPdf(content), new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 2 }
        }).Pages[0];

        Assert.Equal(OfficeColor.Red, OfficeDrawingRasterRenderer.Render(page.ToDrawing()).GetPixel(10, 190));
    }

    [Fact]
    public void AreaRedactionRemovesActualTextFromANoShowTextObject() {
        byte[] bytes = MarkedContentSecurityPdf(
            "BT /Span << /ActualText (secret) >> BDC ET BT /F1 12 Tf 20 100 Td (secret) Tj EMC ET");

        byte[] result = PdfRedactionApplier.Apply(bytes, new[] { new PdfRedactionArea(1, 0, 0, 200, 200) });

        Assert.DoesNotContain("secret", PdfTextExtractor.ExtractAllText(result), System.StringComparison.Ordinal);
        Assert.DoesNotContain("secret", PdfEncoding.Latin1GetString(result), System.StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MatchedGlyphRedactionRefusesToSilentlySkipClippingOnlyText(bool regex) {
        var source = PdfDocument.Load(MarkedContentSecurityPdf("BT /F1 12 Tf 20 100 Td (secret) Tj ET",
            "q BT /F1 12 Tf 7 Tr 20 100 Td (secret) Tj ET Q"));
        var options = new PdfRedactionSearchOptions { TextSelection = PdfRedactionTextSelection.MatchedGlyphs };
        if (regex) options.AddRegex("secret"); else options.AddLiteral("secret");

        PdfRedactionPlan plan = source.Redactions.Search(options);

        Assert.False(plan.IsReviewable);
        Assert.Throws<System.InvalidOperationException>(() => source.Redactions.Apply(plan));
        Assert.Equal(2, PdfTextExtractor.ExtractAllText(source.ToBytes()).Split(new[] { "secret" }, System.StringSplitOptions.None).Length - 1);
    }

    private static byte[] MarkedContentSecurityPdf(string content, string? secondPage = null) {
        var objects = new System.Collections.Generic.List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count " + (secondPage == null ? "1" : "2") + " /Kids [3 0 R " + (secondPage == null ? "" : "6 0 R") + "] /MediaBox [0 0 200 200] /Resources << /Font << /F1 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(5, Encoding.ASCII.GetBytes(content))
        };
        if (secondPage != null) {
            objects.Add("6 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 7 0 R >>\nendobj");
            objects.Add(BuildStreamObject(7, Encoding.ASCII.GetBytes(secondPage)));
        }
        return BuildPdf(objects, rootObjectNumber: 1);
    }
}
