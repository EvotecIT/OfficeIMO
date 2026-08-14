using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_TreatsExplicitNullType3FormGroupAsAbsent() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Form 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Form Do");
        string form = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group null /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, form);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3FormGroup() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Form 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Form Do");
        string form = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group 99 0 R /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, form);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullColorSpaceOnType3ImageMaskAsAbsent() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /ColorSpace null /BitsPerComponent 1", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Mask null")]
    [InlineData("/SMask null")]
    public void RenderPage_TreatsExplicitNullNestedMasksOnType3SoftMaskImageAsAbsent(string nestedMaskEntry) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /SMask 8 0 R", "rgb");
        string softMask = BuildStreamObject(8, $"<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 {nestedMaskEntry}", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image, softMask);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Images);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3ImageDecode() {
        AssertType3FallsBackWithoutNativeShapes(BuildStrictType3ImagePdf("/Decode 99 0 R"));
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3DctDecodeParameters() {
        byte[] jpeg = CreateMinimalJpeg(1, 1);
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /DecodeParms 99 0 R /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedType3PageBlendColorSpace() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        const string graphicsState = "7 0 obj\n<< /Type /ExtGState /BM /Multiply >>\nendobj";
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            "/Group << /Type /Group /S /Transparency /CS 99 0 R >>",
            type3Font,
            glyph,
            graphicsState);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_TreatsExplicitNullType3PageBlendColorSpaceAsOmitted() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        const string graphicsState = "7 0 obj\n<< /Type /ExtGState /BM /Multiply >>\nendobj";
        byte[] pdf = BuildSingleStreamPdfWithPageEntries(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            "/Group << /Type /Group /S /Transparency /CS null >>",
            type3Font,
            glyph,
            graphicsState);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/OCGs 0")]
    [InlineData("/OCGs 99 0 R")]
    public void RenderPage_FailsClosedForMalformedViewUsageApplicationTargets(string targetEntry) {
        const string pageContent = "BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Properties << /Layer 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /OC /Layer BDC 0 0 500 700 re f EMC");
        string content = BuildStreamObject(4, "<<", pageContent);
        string pdfText = string.Join("\n", new[] {
            "%PDF-1.7",
            $"1 0 obj\n<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [7 0 R] /D << /BaseState /ON /AS [<< /Event /View /Category [/View] {targetEntry} >>] >> >> >>\nendobj",
            "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Font << /FType3 5 0 R >> >> /Contents 4 0 R >>\nendobj",
            content,
            type3Font,
            glyph,
            "7 0 obj\n<< /Type /OCG /Name (View target) /Usage << /View << /ViewState /OFF >> >> >>\nendobj",
            "trailer\n<< /Root 1 0 R >>\n%%EOF"
        });

        AssertType3FallsBackWithoutNativeShapes(Encoding.ASCII.GetBytes(pdfText));
    }

    [Fact]
    public void RenderPage_FailsClosedForExifOrientedType3DctImage() {
        byte[] jpeg = CreateType3ReviewJpeg(orientation: 6);
        byte[] pdf = BuildType3ReviewJpegPdf(jpeg, "q 500 0 0 700 0 0 cm /Im1 Do Q");

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_RendersRepeatedValidatedType3DctImagePlacements() {
        byte[] jpeg = CreateType3ReviewJpeg();
        byte[] pdf = BuildType3ReviewJpegPdf(
            jpeg,
            "q 160 0 0 700 0 0 cm /Im1 Do Q q 160 0 0 700 170 0 cm /Im1 Do Q q 160 0 0 700 340 0 cm /Im1 Do Q");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Equal(3, PdfPageImageRenderer.RenderPage(pdf).Images.Count);
        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    private static byte[] CreateType3ReviewJpeg(ushort? orientation = null) {
        var source = new OfficeRasterImage(1, 1, OfficeColor.Red);
        OfficeJpegMetadata metadata = orientation.HasValue
            ? new OfficeJpegMetadata(exif: new byte[] {
                (byte)'I', (byte)'I', 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00,
                0x01, 0x00,
                0x12, 0x01, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00,
                (byte)orientation.Value, (byte)(orientation.Value >> 8), 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00
            })
            : default;
        return OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
            Quality = 100,
            Subsampling = OfficeJpegSubsampling.Y444,
            Metadata = metadata
        });
    }

    private static byte[] BuildType3ReviewJpegPdf(byte[] jpeg, string glyphPaint) {
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", $"500 0 d0 {glyphPaint}");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);
        return pdf;
    }
}
