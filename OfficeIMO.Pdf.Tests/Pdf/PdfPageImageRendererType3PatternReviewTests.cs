using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_PreservesType3ImageInterpolationSelection() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /ImDefault 7 0 R /ImSmooth 8 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 250 0 0 700 0 0 cm /ImDefault Do Q q 250 0 0 700 250 0 cm /ImSmooth Do Q");
        string imageDefault = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "rgb");
        string imageSmooth = BuildStreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Interpolate true", "rgb");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, imageDefault, imageSmooth);

        OfficeDrawingImage[] images = EnumerateImages(PdfPageImageRenderer.RenderPage(pdf)).ToArray();

        Assert.Equal(2, images.Length);
        Assert.False(images[0].Interpolate);
        Assert.True(images[1].Interpolate);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedType3ImageInterpolation() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Interpolate /Bad", "rgb");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForSurplusType3TilingPatternBoxOperands() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5 10] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderedType3TextTracker_DistinguishesCollapsedPaintOrdersByContentPath() {
        var tracker = new RenderedType3TextTracker();
        PdfContentOrderKey rendered = PdfContentOrderKey.Root.Append(2).Append(4).Append(6);
        PdfContentOrderKey ordinary = PdfContentOrderKey.Root.Append(2).Append(4).Append(7);

        tracker.Add(1D, rendered);

        Assert.True(tracker.Contains(1D, rendered));
        Assert.False(tracker.Contains(1D, ordinary));
    }

    [Fact]
    public void RenderPages_ReportsIccApproximationForType3PatternBaseColorSpace() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ColorSpace << /PatternIcc [/Pattern [/ICCBased 8 0 R]] >> /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /PatternIcc cs 0.2 0.4 0.6 /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >>", "0 0 10 10 re f");
        string profile = BuildStreamObject(8, "<< /N 3 /Alternate /DeviceRGB", "fixture-icc-profile");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, profile);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId && diagnostic.Subject == "P1");
    }

    [Fact]
    public void RenderDiagnostics_ChargesType3PatternContentOnce() {
        const string pageContent = "BT /FType3 18 Tf 20 100 Td (A) Tj ET";
        const string glyphContent = "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f";
        string patternContent = "1 0 0 rg 0 0 10 10 re f " + new string(' ', 1024);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", glyphContent);
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >>", patternContent);
        byte[] pdf = BuildSingleStreamPdf(pageContent, "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);
        var readOptions = new PdfReadOptions {
            Limits = new PdfReadLimits {
                MaxPageContentBytes = pageContent.Length + glyphContent.Length + patternContent.Length + 64
            }
        };

        IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics = PdfReadDocument.Open(pdf, readOptions).Pages[0].GetRenderCapabilityDiagnostics();

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderDiagnostics_ChargesNestedType3PatternGlyphsOnce() {
        string outerFont = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string outerGlyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /FInner 8 0 R >> >>", "BT /FInner 8 Tf (A) Tj ET");
        string innerFont = "8 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 9 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string innerGlyph = BuildStreamObject(9, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FOuter 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FOuter 5 0 R >> >>", outerFont, outerGlyph, pattern, innerFont, innerGlyph);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxType3GlyphInvocationsPerPage = 2 }
        });

        IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics = document.Pages[0].GetRenderCapabilityDiagnostics();

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForShearedImageInsideType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 100 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "rgb");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForOptionalContentImageInsideType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /OC 8 0 R", "rgb");
        string optionalContentGroup = "8 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image, optionalContentGroup);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_ColorImageDoesNotConsumeUnusedCallerFillPattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "rgb");
        byte[] pdf = BuildSingleStreamPdf(
            "/Pattern cs /Missing scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyph,
            image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.NotEmpty(drawing.Images);
    }

    [Fact]
    public void RenderPage_FailsClosedForLocallyPatternPaintedImageMaskInColoredType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 8 0 R >> /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn q 500 0 0 700 0 0 cm /Im1 Do Q");
        string imageMask = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x");
        string pattern = BuildStreamObject(8, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "0 1 0 rg 0 0 5 5 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, imageMask, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_SeparatesType3ImageCacheByEffectiveResourceContext() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /F1 7 0 R /F2 8 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /F1 Do q 1 0 0 1 250 0 cm /F2 Do Q");
        string firstForm = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << /ColorSpace << /CS1 [/Indexed /DeviceRGB 0 <FF0000>] >> /XObject << /Im1 9 0 R >> >>", "q 250 0 0 700 0 0 cm /Im1 Do Q");
        string secondForm = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << /ColorSpace << /CS1 [/Indexed /DeviceRGB 0 <00FF00>] >> /XObject << /Im1 9 0 R >> >>", "q 250 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(9, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /CS1 /BitsPerComponent 8 /Filter /ASCIIHexDecode", "00>");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, firstForm, secondForm, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        byte[][] scanlines = EnumerateImages(drawing)
            .Select(static item => PdfPngTestImages.DecodeStoredPngIdat(item.Bytes))
            .ToArray();

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(2, scanlines.Length);
        Assert.Contains(scanlines, bytes => bytes.SequenceEqual(new byte[] { 0, 255, 0, 0 }));
        Assert.Contains(scanlines, bytes => bytes.SequenceEqual(new byte[] { 0, 0, 255, 0 }));
    }

    [Fact]
    public void RenderPage_DecodesType3PatternImagesByEffectiveResourceContext() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /F1 8 0 R /F2 9 0 R >> >>", "/F1 Do q 1 0 0 1 5 0 cm /F2 Do Q");
        string firstForm = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 5 10] /Resources << /ColorSpace << /CS1 [/Indexed /DeviceRGB 0 <FF0000>] >> /XObject << /Im1 10 0 R >> >>", "q 5 0 0 10 0 0 cm /Im1 Do Q");
        string secondForm = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 5 10] /Resources << /ColorSpace << /CS1 [/Indexed /DeviceRGB 0 <00FF00>] >> /XObject << /Im1 10 0 R >> >>", "q 5 0 0 10 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(10, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /CS1 /BitsPerComponent 8 /Filter /ASCIIHexDecode", "00>");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, firstForm, secondForm, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        byte[][] scanlines = EnumerateImages(PdfPageImageRenderer.RenderPage(pdf))
            .Select(static item => PdfPngTestImages.DecodeStoredPngIdat(item.Bytes))
            .ToArray();

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Contains(scanlines, bytes => bytes.SequenceEqual(new byte[] { 0, 255, 0, 0 }));
        Assert.Contains(scanlines, bytes => bytes.SequenceEqual(new byte[] { 0, 0, 255, 0 }));
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedDctImageInsideType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode", "not-a-jpeg");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedDctImageInsideType3Pattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /Im1 8 0 R >> >>", "q 10 0 0 10 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode", "not-a-jpeg");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, image);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData(0xC3)]
    [InlineData(0xC9)]
    public void RenderPage_FailsClosedForUnsupportedDctProcessInsideType3Glyph(int sofMarker) {
        byte[] jpeg = CreateMinimalJpeg(1, 1);
        ReplaceJpegStartOfFrameMarker(jpeg, (byte)sofMarker);
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForType3DctImageWithNonDefaultDecode() {
        byte[] jpeg = CreateMinimalJpeg(1, 1);
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Decode [1 0 1 0 1 0] /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForType3DctImageWithAuthoredColorTransform() {
        byte[] jpeg = CreateMinimalJpeg(1, 1);
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /DecodeParms << /ColorTransform 0 >> /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForType3DctImageThatRequiresIndexedNormalization() {
        byte[] jpeg = CreateMinimalJpeg(1, 1);
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace [/Indexed /DeviceRGB 1 <000000FFFFFF>] /BitsPerComponent 8 /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForType3ShadingPatternWithoutEndpointExtension() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = "7 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> >> >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedMatrixInsideType3ShadingPattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = "7 0 obj\n<< /Type /Pattern /PatternType 2 /Matrix [1 0 0] /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData(" /BBox [0 0 250 700]", " /N 1")]
    [InlineData("", " /N 2")]
    public void RenderPage_FailsClosedForUnhandledType3ShadingSemantics(string shadingEntry, string functionEntry) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = "7 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2" + shadingEntry + " /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1]" + functionEntry + " >> /Extend [true true] >> >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData(" /Domain [0.25 0.75]", "")]
    [InlineData(" /Domain [0 1]", " /Range [0 0.5 0 1 0 1]")]
    public void RenderPage_FailsClosedForUnrepresentedType3ShadingFunctionIntervals(
        string domainEntry,
        string rangeEntry) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = "7 0 obj\n<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2" + domainEntry + " /C0 [1 0 0] /C1 [0 0 1] /N 1" + rangeEntry + " >> /Extend [true true] >> >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderDiagnostics_ReportsUnsupportedDirectShadingInsideType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Shading << /Sh1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Sh1 sh");
        string shading = "7 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, shading);

        IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics = PdfReadDocument.Open(pdf).Pages[0].GetRenderCapabilityDiagnostics();

        Assert.Contains(diagnostics, diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId && diagnostic.Subject == "FType3");
    }

    [Fact]
    public void RenderPage_FailsClosedForNonDefaultStitchedFunctionDomainInsideType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Shading << /Sh1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Sh1 sh");
        string shading = "7 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 3 /Domain [0.25 0.75] /Functions [8 0 R 9 0 R] /Bounds [0.5] /Encode [0 1 0 1] >> /Extend [true true] >>\nendobj";
        string firstFunction = "8 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 1 0] /N 1 >>\nendobj";
        string secondFunction = "9 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, shading, firstFunction, secondFunction);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId && diagnostic.Subject == "FType3");
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingShape);
    }

    [Fact]
    public void RenderPage_FailsClosedForNegativeRadialRadiusInsideType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Shading << /Sh1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Sh1 sh");
        string shading = "7 0 obj\n<< /ShadingType 3 /ColorSpace /DeviceRGB /Coords [0 0 -10 500 700 20] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, shading);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnrepresentableType3ImageTransform() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [1 0 0 1 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 0.001 0 0.5 100 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8", "rgb");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 1 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId && diagnostic.Subject == "FType3");
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingShape);
    }

    [Fact]
    public void RenderPage_FailsClosedForOptionalContentControlledFormInsideStrictType3Pattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /F1 8 0 R >> >>", "/F1 Do");
        string form = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /OC 9 0 R /Resources << >>", "1 0 0 rg 0 0 10 10 re f");
        string optionalContentGroup = "9 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, form, optionalContentGroup);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderDiagnostics_RejectsStitchedType3ShadingWithoutBounds() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Shading << /Sh1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Sh1 sh");
        string shading = "7 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 3 /Domain [0 1] /Functions [8 0 R 9 0 R] /Encode [0 1 0 1] >> /Extend [true true] >>\nendobj";
        string firstFunction = "8 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 1 0] /N 1 >>\nendobj";
        string secondFunction = "9 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 1 0] /C1 [0 0 1] /N 1 >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, shading, firstFunction, secondFunction);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic =>
            diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId && diagnostic.Subject == "FType3");
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingShape);
    }

    [Fact]
    public void RenderPage_FailsClosedForRotatedEllipticalRadialType3Pattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = "7 0 obj\n<< /Type /Pattern /PatternType 2 /Matrix [1.414 1.414 -0.707 0.707 0 0] /Shading << /ShadingType 3 /ColorSpace /DeviceRGB /Coords [250 350 0 250 350 350] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForPartiallyUnderstoodExtGStateInsideType3Pattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /ExtGState << /GS1 8 0 R >> >>", "/GS1 gs 0 0 10 10 re f");
        string graphicsState = "8 0 obj\n<< /Type /ExtGState /ca 0.5 /TR /Identity >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, graphicsState);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnsupportedShadingInsideType3PatternTile() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Shading << /Sh1 8 0 R >> >>", "/Sh1 sh");
        string mesh = BuildStreamObject(8, "<< /ShadingType 4 /ColorSpace /DeviceRGB /BitsPerCoordinate 8 /BitsPerComponent 8 /BitsPerFlag 2 /Decode [0 1 0 1 0 1 0 1 0 1]", "mesh");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, mesh);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnhandledDirectShadingInsideType3PatternTile() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Shading << /Sh1 8 0 R >> >>", "/Sh1 sh");
        string shading = "8 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Domain [0.25 0.75] /Coords [0 0 10 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, shading);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForDirectShadingInsideUncoloredType3PatternTile() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ColorSpace << /PatternRgb [/Pattern /DeviceRGB] >> /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /PatternRgb cs 0 0 1 /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Shading << /Sh1 8 0 R >> >>", "/Sh1 sh");
        string shading = "8 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 10 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, shading);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnknownOperatorInsideType3PatternTile() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >>", "1 0 0 rg 0 0 10 10 re f MadeUp");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedMatrixInsideType3Pattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Matrix [1 0 0] /Resources << >>", "1 0 0 rg 0 0 10 10 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData("/Matrix [1 0 0 1 0]")]
    [InlineData("/BBox [0 0 0 700]")]
    public void RenderPage_FailsClosedForMalformedOrdinaryFormInsideType3Glyph(string formEntry) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Form 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Form Do");
        string boundingBox = formEntry.StartsWith("/BBox", StringComparison.Ordinal) ? string.Empty : "/BBox [0 0 500 700]";
        string form = BuildStreamObject(7, "<< /Type /XObject /Subtype /Form " + boundingBox + " " + formEntry + " /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, form);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedBeforeExpandingOverBudgetType3Pattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 1 1] /XStep 1 /YStep 1 /Resources << >>", "1 0 0 rg 0 0 1 1 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderDrawing_ChargesNestedType3PaintAnalysisToContentDepth() {
        string outerFont = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Font << /FMiddle 7 0 R >> >> >>\nendobj";
        string outerGlyph = BuildStreamObject(6, "<<", "500 0 d0 BT /FMiddle 500 Tf (A) Tj ET");
        string middleFont = "7 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 8 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Font << /FInner 9 0 R >> >> >>\nendobj";
        string middleGlyph = BuildStreamObject(8, "<<", "500 0 d0 BT /FInner 500 Tf (A) Tj ET");
        string innerFont = "9 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 10 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << >> >>\nendobj";
        string innerGlyph = BuildStreamObject(10, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FOuter 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FOuter 5 0 R >> >>", outerFont, outerGlyph, middleFont, middleGlyph, innerFont, innerGlyph);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 1 }
        });

        IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics = document.Pages[0].GetRenderCapabilityDiagnostics();
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenType3PatternTileRecursesIntoActiveGlyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /FType3 5 0 R >> >>", "BT /FType3 8 Tf (A) Tj ET");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForImageMaskInsideUncoloredType3Pattern() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ColorSpace << /PatternRgb [/Pattern /DeviceRGB] >> /Pattern << /P1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 /PatternRgb cs 0 0 1 /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(7, "<< /Type /Pattern /PatternType 1 /PaintType 2 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /Im1 8 0 R >> >>", "q 10 0 0 10 0 0 cm /Im1 Do Q");
        string imageMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, pattern, imageMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForType3DctImageWithUnsupportedColorSpace() {
        byte[] jpeg = CreateMinimalJpeg(1, 1);
        string marker = new string('J', jpeg.Length);
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /Missing /BitsPerComponent 8 /Filter /DCTDecode", marker);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);
        ReplaceAsciiPayload(pdf, marker, jpeg);

        AssertType3FallsBackWithoutNativeShapes(pdf);
        Assert.Contains(Assert.Single(PdfPageImageRenderer.RenderPages(pdf)).CapabilityDiagnostics,
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    private static void ReplaceAsciiPayload(byte[] pdf, string marker, byte[] replacement) {
        byte[] markerBytes = Encoding.ASCII.GetBytes(marker);
        Assert.Equal(markerBytes.Length, replacement.Length);
        for (int offset = 0; offset <= pdf.Length - markerBytes.Length; offset++) {
            bool matches = true;
            for (int index = 0; index < markerBytes.Length; index++) {
                if (pdf[offset + index] == markerBytes[index]) continue;
                matches = false;
                break;
            }
            if (!matches) continue;
            Buffer.BlockCopy(replacement, 0, pdf, offset, replacement.Length);
            return;
        }
        throw new InvalidOperationException("Marker payload was not found.");
    }

    private static IEnumerable<OfficeDrawingImage> EnumerateImages(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingImage image) {
                yield return image;
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                foreach (OfficeDrawingImage nestedImage in EnumerateImages(effectGroup.Drawing)) {
                    yield return nestedImage;
                }
            } else if (element is OfficeDrawingGroup group) {
                foreach (OfficeDrawingImage nestedImage in EnumerateImages(group.Drawing)) {
                    yield return nestedImage;
                }
            } else if (element is OfficeDrawingTilingPattern tilingPattern) {
                foreach (OfficeDrawingImage nestedImage in EnumerateImages(tilingPattern.Tile)) {
                    yield return nestedImage;
                }
            }
        }
    }

    private static void ReplaceJpegStartOfFrameMarker(byte[] jpeg, byte replacement) {
        for (int index = 0; index < jpeg.Length - 1; index++) {
            if (jpeg[index] != 0xFF || jpeg[index + 1] != 0xC0) continue;
            jpeg[index + 1] = replacement;
            return;
        }
        throw new InvalidOperationException("JPEG start-of-frame marker was not found.");
    }
}
