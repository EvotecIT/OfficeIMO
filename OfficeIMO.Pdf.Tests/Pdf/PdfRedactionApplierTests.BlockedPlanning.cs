using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfRedactionApplierTests {
    [Fact]
    public void Apply_RejectsBlockedPlanningOnAnUnselectedPage() {
        byte[] source = BuildBlockedOtherPageRedactionSource();
        byte[] original = source.ToArray();
        var areas = new[] { new PdfRedactionArea(1, 20, 20, 40, 40, "private") };
        PdfDocument document = PdfDocument.Load(source);
        PdfRedactionPlan plan = document.Redactions.Plan(areas);

        Assert.False(plan.IsReviewable);
        Assert.True(plan.Preflight.CanReadLogicalObjects);
        Assert.Contains(plan.Findings, finding => finding.Code == "RedactionSourceTextMappingUnsupported");
        Assert.Single(PdfReadDocument.Open(source).Pages[0].GetImagePlacements());
        Assert.Single(PdfInspector.Inspect(source).GetAnnotationsBySubtype("Text"));
        Assert.Throws<InvalidOperationException>(() => document.Redactions.Apply(areas));
        Assert.Throws<InvalidOperationException>(() => document.Redactions.Apply(plan));
        Assert.Equal(original, source);
        Assert.Equal(original, document.ToBytes());
    }

    [Fact]
    public void Apply_BlockedPlanningPreservesOutputSinks() {
        byte[] source = BuildBlockedOtherPageRedactionSource();
        var areas = new[] { new PdfRedactionArea(1, 20, 20, 40, 40) };
        byte[] sentinel = Encoding.ASCII.GetBytes("existing output");
        using var input = new MemoryStream(source);
        using var output = new MemoryStream();
        output.Write(sentinel, 0, sentinel.Length);
        Assert.Throws<InvalidOperationException>(() => PdfRedactionApplier.Apply(input, areas));
        Assert.Throws<InvalidOperationException>(() => PdfRedactionApplier.Apply(source, output, areas));
        Assert.Equal(sentinel, output.ToArray());
        Assert.Equal(sentinel.Length, output.Position);

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-blocked-redaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string inputPath = Path.Combine(directory, "input.pdf");
            string outputPath = Path.Combine(directory, "output.pdf");
            File.WriteAllBytes(inputPath, source);
            File.WriteAllBytes(outputPath, sentinel);
            Assert.Throws<InvalidOperationException>(() => PdfRedactionApplier.Apply(inputPath, outputPath, areas));
            Assert.Equal(source, File.ReadAllBytes(inputPath));
            Assert.Equal(sentinel, File.ReadAllBytes(outputPath));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void RemoveImagePlacements_RejectsBlockedMatchInspection() {
        byte[] source = BuildBlockedOtherPageRedactionSource();
        PdfImagePlacement placement = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetImagePlacements(1));

        Assert.Throws<InvalidOperationException>(() => PdfRedactionApplier.RemoveImagePlacements(source, new[] { placement }));
    }

    [Fact]
    public void RemoveTextInAreas_PreservesItsIndependentTextOnlyContract() {
        byte[] source = BuildBlockedOtherPageRedactionSource();
        var areas = new[] { new PdfRedactionArea(1, 70, 95, 130, 20) };

        byte[] result = PdfRedactionApplier.RemoveTextInAreas(source, areas);

        Assert.DoesNotContain(PdfReadDocument.Open(result).Pages[0].GetTextSpans(), span => span.Text.Contains("Known body"));
        Assert.Single(PdfReadDocument.Open(result).Pages[0].GetImagePlacements());
        Assert.Single(PdfInspector.Inspect(result).GetAnnotationsBySubtype("Text"));
        Assert.Contains("Sensitive annotation", PdfEncoding.Latin1GetString(result), StringComparison.Ordinal);
        Assert.Contains("<3042> Tj", PdfEncoding.Latin1GetString(result), StringComparison.Ordinal);
    }

    private static byte[] BuildBlockedOtherPageRedactionSource() {
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] /MediaBox [0 0 300 300] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 7 0 R >> >> /Contents 6 0 R /Annots [8 0 R] >>\nendobj",
            "4 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F2 9 0 R >> >> /Contents 11 0 R >>\nendobj",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj",
            BuildStreamObject(6, Encoding.ASCII.GetBytes("q 40 0 0 40 20 20 cm /Im1 Do Q\nBT /F1 12 Tf 72 100 Td (Known body) Tj ET")),
            BuildStreamObject(7, Encoding.ASCII.GetBytes("80>"), "/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /ASCIIHexDecode"),
            "8 0 obj\n<< /Type /Annot /Subtype /Text /Rect [20 20 60 60] /Contents (Sensitive annotation) >>\nendobj",
            "9 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiMin-W3 /Encoding /UnsupportedTestMap /DescendantFonts [10 0 R] >>\nendobj",
            "10 0 obj\n<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiMin-W3 /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 5 >> /DW 1000 >>\nendobj",
            BuildStreamObject(11, Encoding.ASCII.GetBytes("BT /F2 12 Tf 72 100 Td <3042> Tj ET"))
        };
        return BuildPdf(objects, rootObjectNumber: 1);
    }
}
