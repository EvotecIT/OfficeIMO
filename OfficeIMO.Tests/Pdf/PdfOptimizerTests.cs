using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOptimizerTests {
    [Fact]
    public void Optimize_CompressesUnfilteredStreamsAndPreservesText() {
        byte[] source = BuildPdfWithUncompressedTextStream("BT\n/F1 12 Tf\n72 720 Td\n(" + new string('A', 4096) + ") Tj\nET\n");

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source);

        Assert.True(result.Applied);
        Assert.True(result.OptimizedLengthBytes < result.OriginalLengthBytes);
        PdfOptimizationAction action = Assert.Single(result.Actions);
        Assert.Equal("CompressStream", action.Kind);
        Assert.Equal(5, action.ObjectNumber);
        Assert.Contains("/Filter /FlateDecode", PdfEncoding.Latin1GetString(result.Bytes), StringComparison.Ordinal);
        Assert.Contains(new string('A', 64), PdfTextExtractor.ExtractAllText(result.Bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void Optimize_ReturnsOriginalWhenCandidateIsNotSmaller() {
        byte[] source = BuildPdfWithUncompressedTextStream("BT\n/F1 12 Tf\n72 720 Td\n(Tiny) Tj\nET\n");

        PdfOptimizationActionResult result = PdfOptimizer.Optimize(source);

        Assert.False(result.Applied);
        Assert.True(result.ReturnedOriginal);
        Assert.Equal(source.Length, result.Bytes.Length);
        Assert.Equal(source, result.Bytes);
    }

    [Fact]
    public void Optimize_RejectsSignedPdf() {
        byte[] signed = Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 3 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "3 0 obj",
            "<< /Fields [4 0 R] /SigFlags 1 >>",
            "endobj",
            "4 0 obj",
            "<< /FT /Sig /V 5 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Sig /ByteRange [0 0 0 0] /Contents <> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        }));

        Assert.Throws<NotSupportedException>(() => PdfOptimizer.Optimize(signed));
    }

    private static byte[] BuildPdfWithUncompressedTextStream(string streamContent) {
        byte[] streamBytes = Encoding.ASCII.GetBytes(streamContent);
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + streamBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            streamContent.TrimEnd('\n'),
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "startxref",
            "123",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
