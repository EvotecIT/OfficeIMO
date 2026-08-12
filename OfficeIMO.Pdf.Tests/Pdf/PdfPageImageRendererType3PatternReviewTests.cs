using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
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
    public void RenderPage_FailsClosedForMalformedDctImageInsideType3Glyph() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /XObject << /Im1 7 0 R >> >> >>\nendobj";
        string glyph = BuildStreamObject(6, "<<", "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(7, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode", "not-a-jpeg");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyph, image);

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
}
