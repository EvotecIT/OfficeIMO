using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void ExistingVisualAnnotations_FlattenPlacesAppearanceStreamsUsingBBoxAndPreservesMatrixResources() {
        byte[] annotated = BuildAppearancePlacementAnnotationPdf();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.Contains("/BBox [10 20 150 60]", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Matrix [140 0 0 40 0 0]", beforePdf, StringComparison.Ordinal);
        Assert.Equal(2, PdfInspector.Inspect(annotated).AnnotationCount);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Annots [", pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot2 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 10 60 cm", pdf, StringComparison.Ordinal);
        Assert.Contains("1 0 0 1 20 140 cm", pdf, StringComparison.Ordinal);
        Assert.Contains("/Matrix [ 140 0 0 40 0 0 ]", pdf, StringComparison.Ordinal);
        Assert.Contains("/Resources << /Font << /F1 ", pdf, StringComparison.Ordinal);
        Assert.Contains("(BBox AP) Tj", pdf, StringComparison.Ordinal);
        Assert.Contains("(Matrix AP) Tj", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenMaterializesInheritedPageResources() {
        byte[] annotated = BuildInheritedResourceAnnotationPdf();
        string beforePdf = Encoding.ASCII.GetString(annotated);
        Assert.Contains("/Resources << /Font << /F1 7 0 R >> >>", beforePdf, StringComparison.Ordinal);
        Assert.Contains("/Contents 4 0 R /Annots [5 0 R]", beforePdf, StringComparison.Ordinal);

        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);

        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.DoesNotContain("/Subtype /FreeText", pdf, StringComparison.Ordinal);
        Assert.Contains("(Inherited Text) Tj", pdf, StringComparison.Ordinal);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Matches(@"/Resources << /Font << /F1 \d+ 0 R >> /XObject << /OfficeIMOAnnot1 \d+ 0 R >> >>", pdf);
    }

    [Fact]
    public void ExistingVisualAnnotations_FlattenDoesNotMutateSharedPageResourcesForUntouchedPages() {
        byte[] annotated = BuildSharedResourceTwoPageAnnotationPdf();
        byte[] flattened = PdfAnnotationFlattener.FlattenVisualAnnotations(annotated);
        string pdf = Encoding.ASCII.GetString(flattened);
        var (objects, _) = PdfSyntax.ParseObjects(flattened);
        List<PdfDictionary> pages = objects
            .OrderBy(pair => pair.Key)
            .Select(pair => pair.Value.Value)
            .OfType<PdfDictionary>()
            .Where(dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page")
            .ToList();

        Assert.Equal(2, pages.Count);
        Assert.Equal(0, PdfInspector.Inspect(flattened).AnnotationCount);
        Assert.Contains("/OfficeIMOAnnot1 Do", pdf, StringComparison.Ordinal);
        Assert.Contains("(Shared resource page two) Tj", pdf, StringComparison.Ordinal);

        PdfDictionary firstResources = Assert.IsType<PdfDictionary>(pages[0].Items["Resources"]);
        PdfDictionary firstXObjects = Assert.IsType<PdfDictionary>(firstResources.Items["XObject"]);
        Assert.True(firstXObjects.Items.ContainsKey("OfficeIMOAnnot1"));

        PdfReference secondResourcesReference = Assert.IsType<PdfReference>(pages[1].Items["Resources"]);
        PdfDictionary secondResources = Assert.IsType<PdfDictionary>(objects[secondResourcesReference.ObjectNumber].Value);
        Assert.True(secondResources.Items.ContainsKey("Font"));
        Assert.False(secondResources.Items.ContainsKey("XObject"));
    }

    private static byte[] BuildAppearancePlacementAnnotationPdf() {
        const string pageContent = "";
        const string bboxAppearance = "BT /F1 12 Tf 10 20 Td (BBox AP) Tj ET";
        const string matrixAppearance = "BT /F1 12 Tf 0 0 Td (Matrix AP) Tj ET";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int bboxAppearanceLength = Encoding.ASCII.GetByteCount(bboxAppearance);
        int matrixAppearanceLength = Encoding.ASCII.GetByteCount(matrixAppearance);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 8 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 80 160 120] /Contents (BBox appearance) /AP << /N 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [10 20 150 60] /Resources << /Font << /F1 7 0 R >> >> /Length {bboxAppearanceLength} >>",
            "stream",
            bboxAppearance,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 140 160 180] /Contents (Matrix appearance) /AP << /N 9 0 R >> >>",
            "endobj",
            "9 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 1 1] /Matrix [140 0 0 40 0 0] /Resources << /Font << /F1 7 0 R >> >> /Length {matrixAppearanceLength} >>",
            "stream",
            matrixAppearance,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildInheritedResourceAnnotationPdf() {
        const string pageContent = "BT /F1 12 Tf 20 40 Td (Inherited Text) Tj ET";
        const string appearance = "BT /F1 10 Tf 0 0 Td (Inherited Resource AP) Tj ET";
        int pageStreamLength = Encoding.ASCII.GetByteCount(pageContent);
        int appearanceLength = Encoding.ASCII.GetByteCount(appearance);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] /Resources << /Font << /F1 7 0 R >> >> >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Length {pageStreamLength} >>",
            "stream",
            pageContent,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 80 120 110] /Contents (Inherited resource annotation) /AP << /N 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 100 30] /Resources << /Font << /F1 7 0 R >> >> /Length {appearanceLength} >>",
            "stream",
            appearance,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSharedResourceTwoPageAnnotationPdf() {
        const string firstPageContent = "BT /F1 12 Tf 20 40 Td (Shared resource page one) Tj ET";
        const string secondPageContent = "BT /F1 12 Tf 20 40 Td (Shared resource page two) Tj ET";
        const string appearance = "BT /F1 10 Tf 0 0 Td (Shared resource AP) Tj ET";
        int firstPageStreamLength = Encoding.ASCII.GetByteCount(firstPageContent);
        int secondPageStreamLength = Encoding.ASCII.GetByteCount(secondPageContent);
        int appearanceLength = Encoding.ASCII.GetByteCount(appearance);

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 5 0 R /Resources 8 0 R /Annots [9 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 6 0 R /Resources 8 0 R >>",
            "endobj",
            "5 0 obj",
            $"<< /Length {firstPageStreamLength} >>",
            "stream",
            firstPageContent,
            "endstream",
            "endobj",
            "6 0 obj",
            $"<< /Length {secondPageStreamLength} >>",
            "stream",
            secondPageContent,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            "endobj",
            "8 0 obj",
            "<< /Font << /F1 7 0 R >> >>",
            "endobj",
            "9 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [20 90 140 120] /Contents (Shared resource annotation) /AP << /N 10 0 R >> >>",
            "endobj",
            "10 0 obj",
            $"<< /Type /XObject /Subtype /Form /BBox [0 0 120 30] /Resources << /Font << /F1 7 0 R >> >> /Length {appearanceLength} >>",
            "stream",
            appearance,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
