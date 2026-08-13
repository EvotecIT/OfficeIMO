using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
    [Fact]
    public void RenderPage_ChargesPaintChannelFormsToSharedPageContentBudget() {
        const string pageContent = "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        const string glyphContent = "500 0 d0 /F1 Do /F2 Do";
        string firstFormContent = new string(' ', 128) + "10 10 1 1 re f";
        string secondFormContent = new string(' ', 128) + "20 20 1 1 re f";
        byte[] pdf = BuildPaintChannelBudgetPdf(
            pageContent,
            glyphContent,
            firstFormContent,
            secondFormContent);
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxPageContentBytes = pageContent.Length + glyphContent.Length + firstFormContent.Length + 8
            }
        };

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfPageImageRenderer.RenderPages(
                pdf,
                options: new PdfPageRenderOptions { ContinueOnError = false },
                readOptions: readOptions));

        Assert.Equal(PdfReadLimitKind.PageContentBytes, exception.Kind);
    }

    [Fact]
    public void RenderPage_SeparatesSharedGlyphPaintChannelsByInheritedResources() {
        byte[] pdf = BuildInheritedPaintChannelResourcePdf();

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(22, 96));
        OfficeColor stroke = raster.GetPixel(81, 96);
        Assert.Equal((byte)0, stroke.R);
        Assert.Equal((byte)0, stroke.G);
        Assert.Equal((byte)255, stroke.B);
        Assert.True(stroke.A > 0);
    }

    [Fact]
    public void RenderPage_AnalyzesGlyphFormsInTheProjectedTextPosition() {
        const string pageContent = "/OC /Hidden BDC /Pattern cs /P1 scn EMC " +
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [10 0 R] /D << /OFF [10 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> /Pattern << /P1 9 0 R >> /Properties << /Hidden 10 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", pageContent),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Nested 8 0 R >> >> >>\nendobj",
            StreamObject(6, "<<", "500 0 d0 /Nested Do"),
            StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [300 0 500 700]", "300 0 200 700 re f"),
            StreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f"),
            "10 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj"
        };
        byte[] pdf = Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(27, 96));
    }

    [Fact]
    public void RenderPage_IgnoresPatternConsumptionOutsideTransformedImageQuad() {
        const string pageContent = "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        const string groupContent = "q /OC /Hidden BDC /Pattern cs /Bad scn EMC " +
            "100 0 10 10 re W n 100 100 -100 100 200 0 cm /Im1 Do Q 0 0 500 700 re f";
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [11 0 R] /D << /OFF [11 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> /Pattern << /P1 7 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", pageContent),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Group 8 0 R >> >> >>\nendobj",
            StreamObject(6, "<<", "500 0 d0 /Group Do"),
            StreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f"),
            StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /K false /CS /DeviceRGB >> /Resources << /Pattern << /Bad 9 0 R >> /XObject << /Im1 10 0 R >> /Properties << /Hidden 11 0 R >> >>", groupContent),
            StreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 0 1 rg 0 0 5 5 re f"),
            StreamObject(10, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /ASCIIHexDecode", "FF0000>"),
            "11 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj"
        };
        byte[] pdf = Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Empty(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.NotEmpty(drawing.Elements);
    }

    [Fact]
    public void RenderPage_ChargesNestedGlyphPaintAnalysisToSharedBudget() {
        const string pageContent = "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        const string glyphContent = "500 0 d0 q /OC /Hidden BDC /Pattern cs /Bad scn EMC /Fm1 Do Q 0 0 500 700 re f";
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [12 0 R] /D << /OFF [12 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> /Pattern << /P1 10 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", pageContent),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /Bad 11 0 R >> /XObject << /Fm1 7 0 R >> /Properties << /Hidden 12 0 R >> >> >>\nendobj",
            StreamObject(6, "<<", glyphContent),
            StreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << /Font << /Nested 8 0 R >> >>", "BT /Nested 500 Tf (BB) Tj ET"),
            "8 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 9 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << >> >>\nendobj",
            StreamObject(9, "<<", "500 0 d0 0 0 500 700 re f"),
            StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f"),
            StreamObject(11, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 0 1 rg 0 0 5 5 re f"),
            "12 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj"
        };
        byte[] pdf = Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxType3GlyphInvocationsPerPage = 2 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.Type3GlyphInvocations, exception.Kind);
        Assert.Equal(2, exception.Limit);
        Assert.Equal(3, exception.Actual);
    }

    private static byte[] BuildPaintChannelBudgetPdf(
        string pageContent,
        string glyphContent,
        string firstFormContent,
        string secondFormContent) {
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> /Pattern << /P1 10 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", pageContent),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /F1 8 0 R /F2 9 0 R >> >> >>\nendobj",
            StreamObject(6, "<<", glyphContent),
            StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100]", firstFormContent),
            StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100]", secondFormContent),
            StreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f")
        };
        return Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
    }

    private static byte[] BuildInheritedPaintChannelResourcePdf() {
        const string pageContent = "/OC /Hidden BDC /Pattern cs /PFill scn /Pattern CS /PStroke SCN EMC " +
            "BT /F1 18 Tf 20 100 Td (A) Tj ET BT /F2 18 Tf 80 100 Td (A) Tj ET";
        string[] objects = {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [13 0 R] /D << /OFF [13 0 R] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R /F2 6 0 R >> /Pattern << /PFill 11 0 R /PStroke 12 0 R >> /Properties << /Hidden 13 0 R >> >> /Contents 4 0 R >>\nendobj",
            StreamObject(4, "<<", pageContent),
            "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 7 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Nested 8 0 R >> >> >>\nendobj",
            "6 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 2 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 7 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Nested 9 0 R >> >> >>\nendobj",
            StreamObject(7, "<<", "500 0 d0 /Nested Do"),
            StreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700]", "0 0 500 700 re f"),
            StreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700]", "1 j 60 w 30 30 440 640 re S"),
            StreamObject(11, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f"),
            StreamObject(12, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 0 1 rg 0 0 5 5 re f"),
            "13 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj"
        };
        return Encoding.ASCII.GetBytes("%PDF-1.4\n" + string.Join("\n", objects) + "\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
    }
}
