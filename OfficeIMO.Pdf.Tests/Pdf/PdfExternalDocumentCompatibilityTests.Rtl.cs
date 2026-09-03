using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfExternalDocumentCompatibilityTests {
    [Fact]
    public void ExternalToUnicodeLogicalRtlWithRightToLeftProgressionPreservesOrder() {
        byte[] bytes = BuildExternalRtlToUnicodePdf(
            new[] { "0627", "0644", "0639", "0631", "0628", "064A", "0629" },
            "-12 0 0 12 500 720");

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.False(span.GlyphSequenceProgressesLeftToRight);
        Assert.Equal("العربية", span.Text);
        Assert.Contains("العربية", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void ExternalToUnicodeVisualRtlWithLeftToRightProgressionRestoresLogicalOrder() {
        byte[] bytes = BuildExternalRtlToUnicodePdf(
            new[] { "0629", "064A", "0628", "0631", "0639", "0644", "0627" },
            "12 0 0 12 72 720");

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetTextSpans());
        string extracted = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(span.GlyphSequenceProgressesLeftToRight);
        Assert.Equal("ةيبرعلا", span.Text);
        Assert.Contains("العربية", extracted, StringComparison.Ordinal);
    }

    private static byte[] BuildExternalRtlToUnicodePdf(
        IReadOnlyList<string> unicodeHexValues,
        string textMatrix) {
        string glyphHex = string.Concat(Enumerable.Range(1, unicodeHexValues.Count)
            .Select(static value => value.ToString("X4", System.Globalization.CultureInfo.InvariantCulture)));
        var cmap = new StringBuilder()
            .AppendLine("/CIDInit /ProcSet findresource begin")
            .AppendLine("12 dict begin")
            .AppendLine("begincmap")
            .AppendLine("1 begincodespacerange")
            .AppendLine("<0000> <FFFF>")
            .AppendLine("endcodespacerange")
            .AppendLine(unicodeHexValues.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) + " beginbfchar");
        for (int index = 0; index < unicodeHexValues.Count; index++) {
            cmap.Append('<')
                .Append((index + 1).ToString("X4", System.Globalization.CultureInfo.InvariantCulture))
                .Append("> <")
                .Append(unicodeHexValues[index])
                .AppendLine(">");
        }
        cmap.AppendLine("endbfchar")
            .AppendLine("endcmap")
            .AppendLine("CMapName currentdict /CMap defineresource pop")
            .AppendLine("end")
            .AppendLine("end");

        byte[] content = Encoding.ASCII.GetBytes(
            $"BT\n/F13 1 Tf\n{textMatrix} Tm\n<{glyphHex}> Tj\nET\n");
        var objects = new[] {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 612 792] /Resources << /Font << /F13 4 0 R >> >> >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 5 0 R >>\nendobj",
            "4 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /ExternalRtl /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 6 0 R >>\nendobj",
            BuildStreamObject(5, content),
            BuildStreamObject(6, Encoding.ASCII.GetBytes(cmap.ToString())),
            "7 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /ExternalRtl /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /DW 1000 >>\nendobj"
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }
}
