using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void ExistingVisualAnnotations_FlattenSynthesizesInkPolygonAndPolylineWithoutAppearance() {
        byte[] annotated = BuildPathAnnotationsPdfWithoutAppearances();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.DoesNotContain("/AP", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Ink", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /Polygon", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Subtype /PolyLine", beforePdf, StringComparison.Ordinal);
        Assert.Equal(3, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Ink", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Polygon", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /PolyLine", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot3 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("0.2 0.3 0.8 RG 3 w 1 J 1 j", pdf, StringComparison.Ordinal);
        Assert.Contains("10 10 m 40 40 l 80 20 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("90 10 m 110 35 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0.9 0.6 rg 0.8 0.1 0.1 RG 2 w 10 10 m 60 50 l 110 10 l h B", pdf, StringComparison.Ordinal);
        Assert.Contains("0.1 0.6 0.2 RG 2 w 10 10 m 60 10 l 110 10 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("18 6.4 m 10 10 l 18 13.6 l S", pdf, StringComparison.Ordinal);
        Assert.Contains("0.9 0.7 0.2 rg 110 10 m 102 13.6 l 102 6.4 l h B", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenRemovesDependentPopupAnnotations() {
        byte[] annotated = BuildFreeTextAnnotationWithPopupPdf();

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /Popup", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Parent 5 0 R", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenCompensatesAppearanceMatrixPlacement() {
        byte[] annotated = BuildFreeTextAnnotationWithAppearanceMatrixPdf();

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.Contains("1 0 0 1 3 7 cm", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("1 0 0 1 10 20 cm", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenSkipsNoViewAnnotations() {
        byte[] annotated = BuildVisibleAndNoViewFreeTextAnnotationsPdf();

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(1, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Contents (Visible note)", pdf, StringComparison.Ordinal);
        Assert.Contains("/Contents (Hidden note)", pdf, StringComparison.Ordinal);
        Assert.Contains("/F 32", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
    }

    private static byte[] BuildPathAnnotationsPdfWithoutAppearances() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R 7 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Ink /Rect [20 80 160 130] /Contents (Synthetic ink) /InkList [[30 90 60 120 100 100] [110 90 130 115]] /C [0.2 0.3 0.8] /Border [0 0 3] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Polygon /Rect [20 140 160 200] /Contents (Synthetic polygon) /Vertices [30 150 80 190 130 150] /C [0.8 0.1 0.1] /IC [1 0.9 0.6] /Border [0 0 2] >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /PolyLine /Rect [20 210 160 260] /Contents (Synthetic polyline) /Vertices [30 220 80 220 130 220] /C [0.1 0.6 0.2] /IC [0.9 0.7 0.2] /Border [0 0 2] /LE [/OpenArrow /ClosedArrow] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextAnnotationWithPopupPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 7 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 40 120 80] /Contents (Flatten me) /AP << /N 6 0 R >> /Popup 7 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Length 8 >>",
            "stream",
            "0 0 m S",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /Popup /Rect [120 80 180 140] /Parent 5 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextAnnotationWithAppearanceMatrixPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 110 60] /Contents (Matrix) /AP << /N 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 50 20] /Matrix [2 0 0 2 4 6] /Length 8 >>",
            "stream",
            "0 0 m S",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildVisibleAndNoViewFreeTextAnnotationsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 7 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 40 120 80] /Contents (Visible note) /AP << /N 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Length 8 >>",
            "stream",
            "0 0 m S",
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 100 120 140] /Contents (Hidden note) /F 32 /AP << /N 8 0 R >> >>",
            "endobj",
            "8 0 obj",
            "<< /Type /XObject /Subtype /Form /BBox [0 0 100 40] /Length 8 >>",
            "stream",
            "0 0 m S",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
