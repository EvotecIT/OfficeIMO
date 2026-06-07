using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesStampAndCaretWithoutAppearance() {
        byte[] annotated = BuildStampAndCaretAnnotationsPdfWithoutAppearances();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Stamp", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Caret", beforePdf, StringComparison.Ordinal);
        Assert.Equal(2, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Stamp", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Caret", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Helvetica", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.95 0.9 rg 0 0 140 40 re f", pdf, StringComparison.Ordinal);
        Assert.Contains("0.8 0.05 0.05 RG 2 w 1 1 138 38 re S", pdf, StringComparison.Ordinal);
        Assert.Contains("BT /Helv 15.2 Tf 0.8 0.05 0.05 rg 33.9 12.4 Td <544F5020534543524554> Tj ET", pdf, StringComparison.Ordinal);
        Assert.Contains("0.1 0.1 0.1 RG 2 w 1 J 1 j 0 30 m 10 0 l 20 30 l S", pdf, StringComparison.Ordinal);
    }

    private static byte[] BuildStampAndCaretAnnotationsPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Stamp /Rect [20 80 160 120] /Contents (Synthetic stamp) /Name /TopSecret /C [0.8 0.05 0.05] /IC [1 0.95 0.9] /Border [0 0 2] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Caret /Rect [30 150 50 180] /Contents (Synthetic caret) /C [0.1 0.1 0.1] /Border [0 0 2] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
