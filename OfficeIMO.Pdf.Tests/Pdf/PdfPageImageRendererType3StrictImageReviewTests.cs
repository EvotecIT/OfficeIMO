using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_ChargesAnnotationAppearanceImageTintsToSharedPageBudget() {
        const string firstSeparation = "[/Separation /Spot1 /DeviceRGB 9 0 R]";
        const string secondSeparation = "[/Separation /Spot2 /DeviceRGB 10 0 R]";
        string annotation = "5 0 obj\n<< /Type /Annot /Subtype /Stamp /Rect [10 10 50 50] /AP << /N 6 0 R >> >>\nendobj";
        string appearance = BuildStreamObject(
            6,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 40 40] " +
            "/Resources << /XObject << /Im1 7 0 R /Im2 8 0 R >> >>",
            "q 10 0 0 10 0 0 cm /Im1 Do Q q 10 0 0 10 20 0 cm /Im2 Do Q");
        string firstImage = BuildStreamObject(
            7,
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 4 /BitsPerComponent 8 /ColorSpace " + firstSeparation,
            new string('@', 16));
        string secondImage = BuildStreamObject(
            8,
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 4 /BitsPerComponent 8 /ColorSpace " + secondSeparation,
            new string('#', 16));
        string firstTint = BuildStreamObject(9, "<< /FunctionType 4 /Domain [0 1] /Range [0 1 0 1 0 1]", "{ dup dup }");
        string secondTint = BuildStreamObject(10, "<< /FunctionType 4 /Domain [0 1] /Range [0 1 0 1 0 1]", "{ dup dup }");
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            string.Empty,
            "<< >>",
            "/Annots [5 0 R]",
            annotation,
            appearance,
            firstImage,
            secondImage,
            firstTint,
            secondTint);
        PdfReadPage page = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentOperations = 50 }
        }).Pages[0];

        OfficeDrawing drawing = page.ToDrawing();

        int normalizedImages = 0;
        foreach (OfficeDrawingImage image in drawing.Images) {
            if (OfficePngReader.TryDecode(image.Bytes, out _)) normalizedImages++;
        }
        Assert.Equal(1, normalizedImages);
    }

    [Fact]
    public void RenderPage_ReusesDrawingPassBudgetForTopLevelImageExtraction() {
        const string firstSeparation = "[/Separation /Spot1 /DeviceRGB 7 0 R]";
        const string secondSeparation = "[/Separation /Spot2 /DeviceRGB 8 0 R]";
        string firstImage = BuildStreamObject(
            5,
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 4 /BitsPerComponent 8 /ColorSpace " + firstSeparation,
            new string('@', 16));
        string secondImage = BuildStreamObject(
            6,
            "<< /Type /XObject /Subtype /Image /Width 4 /Height 4 /BitsPerComponent 8 /ColorSpace " + secondSeparation,
            new string('#', 16));
        string firstTint = BuildStreamObject(7, "<< /FunctionType 4 /Domain [0 1] /Range [0 1 0 1 0 1]", "{ dup dup }");
        string secondTint = BuildStreamObject(8, "<< /FunctionType 4 /Domain [0 1] /Range [0 1 0 1 0 1]", "{ dup dup }");
        byte[] pdf = BuildSingleStreamPdf(
            "q 10 0 0 10 0 0 cm /Im1 Do Q q 10 0 0 10 20 0 cm /Im2 Do Q",
            "<< /XObject << /Im1 5 0 R /Im2 6 0 R >> >>",
            firstImage,
            secondImage,
            firstTint,
            secondTint);
        PdfReadPage page = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxContentOperations = 50 }
        }).Pages[0];

        OfficeDrawing drawing = page.ToDrawing();

        int normalizedImages = 0;
        foreach (OfficeDrawingImage image in drawing.Images) {
            if (OfficePngReader.TryDecode(image.Bytes, out _)) normalizedImages++;
        }
        Assert.Equal(1, normalizedImages);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenViewUsageTargetsUndeclaredOcg() {
        const string pageContent = "BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Properties << /Layer 8 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /OC /Layer BDC 0 0 500 700 re f EMC");
        string content = BuildStreamObject(4, "<<", pageContent);
        string pdfText = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [7 0 R] /D << /BaseState /ON /AS [<< /Event /View /Category [/View] /OCGs [8 0 R] >>] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            content,
            type3Font,
            glyph,
            "7 0 obj\n<< /Type /OCG /Name (Declared) >>\nendobj",
            "8 0 obj\n<< /Type /OCG /Name (Undeclared) /Usage << /View << /ViewState /OFF >> >> >>\nendobj",
            "trailer\n<< /Root 1 0 R >>\n%%EOF"
        });

        AssertType3FallsBackWithoutNativeShapes(Encoding.ASCII.GetBytes(pdfText));
    }

    [Theory]
    [InlineData("/SMask null")]
    [InlineData("/Mask null")]
    public void RenderPage_TreatsExplicitNullType3ImageTransparencyMaskAsAbsent(string maskEntry) {
        byte[] pdf = BuildStrictType3ImagePdf(maskEntry);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Filter 0")]
    [InlineData("/Filter []")]
    [InlineData("/Filter [/FlateDecode 9 0 R]")]
    public void RenderPage_FailsClosedForMalformedType3ImageFilter(string filterEntry) {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf(filterEntry));
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3ImageOptionalContent() {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf("/OC 99 0 R"));
    }

    [Theory]
    [InlineData("/Perceptual ri", "")]
    [InlineData("", "/Intent /Perceptual")]
    public void RenderPage_PreservesSupportedAuthoredType3ImageRenderingIntent(string glyphPrefix, string imageEntry) {
        byte[] pdf = BuildStrictType3ImagePdf(imageEntry, glyphPrefix);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(
            result.CapabilityDiagnostics,
            static diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Intent /Bogus")]
    [InlineData("/Intent 0")]
    [InlineData("/Intent 99 0 R")]
    public void RenderPage_FailsClosedForMalformedType3ImageRenderingIntent(string imageEntry) {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf(imageEntry));
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullType3TransparencyGroupKnockoutAsDefaultFalse() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Group 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Group Do");
        string group = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /Type /Group /S /Transparency /I true /K null /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyph,
            group);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullType3FormOptionalContentAsAbsent() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Form 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Form Do");
        string form = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /OC null /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyph,
            form);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    private static byte[] BuildStrictType3ImagePdf(string imageEntry, string glyphPrefix = "") {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", $"500 0 d0 {glyphPrefix} q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 {imageEntry}", "rgb");
        return BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyph,
            image);
    }
}
