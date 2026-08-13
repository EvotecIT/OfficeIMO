using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void RenderPage_FailsClosedWhenType3SoftMaskUsesInheritedStrokeState() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 25 w /GS1 gs 0 0 m 500 700 l S");
        string state = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string mask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << >>", "0 0 m 500 700 l S");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, state, mask);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingShape);
    }

    [Fact]
    public void RenderPage_CapturesType3SoftMaskStrokeStateAtInstallation() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 2 w /GS1 gs 1 w 0 0 m 500 700 l S");
        string state = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string mask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << >>", "0 0 m 500 700 l S");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, state, mask);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.DoesNotContain(drawing.Elements, element => element is OfficeDrawingShape);
    }

    [Fact]
    public void RenderPage_DoesNotReuseLuminosityMaskValidationAcrossInheritedColors() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 q /GS1 gs 0 0 100 700 re f Q 1 0 0 rg /GS1 gs 200 0 100 700 re f");
        string state = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string mask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, state, mask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForOptionalContentOnNestedFormInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /XObject << /Fm1 9 0 R >> >>", "/Fm1 Do");
        string nestedForm = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /OC 10 0 R /Resources << >>", "0 0 500 700 re f");
        string optionalContent = "10 0 obj\n<< /Type /OCG /Name (Nested layer) >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, nestedForm, optionalContent);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData("/Matrix [1 0 0 1 0]")]
    [InlineData("/BBox [0 0 0 700]")]
    public void RenderPage_FailsClosedForMalformedNestedFormInsideType3SoftMask(string formEntry) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /XObject << /Fm1 9 0 R >> >>", "/Fm1 Do");
        string boundingBox = formEntry.StartsWith("/BBox", StringComparison.Ordinal) ? string.Empty : "/BBox [0 0 500 700]";
        string nestedForm = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form " + boundingBox + " " + formEntry + " /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, nestedForm);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForMalformedOperatorInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << >>", "/Bad 0 0 1 0 0 cm 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData("/Pattern cs /Missing scn 0 0 500 700 re f")]
    [InlineData("/Pattern CS /Missing SCN 10 w 0 0 m 500 700 l S")]
    public void RenderPage_FailsClosedForMissingDirectPatternInsideType3SoftMask(string maskContent) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << >> >>", maskContent);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData("/Missing cs 0 0 0 sc 0 0 500 700 re f")]
    [InlineData("/Missing CS 0 0 0 SC 10 w 0 0 m 500 700 l S")]
    public void RenderPage_FailsClosedForMissingDirectColorSpaceInsideType3SoftMask(string maskContent) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << >>", maskContent);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData("q /GSZero gs /Bad Do Q 0 0 500 700 re f", "/ExtGState << /GSZero 10 0 R >>", "10 0 obj\n<< /Type /ExtGState /ca 0 >>\nendobj")]
    [InlineData("q 600 600 10 10 re W n /Bad Do Q 0 0 500 700 re f", "", null)]
    public void RenderPage_IgnoresInvisibleTransparencyFormInsideType3SoftMask(string maskContent, string extraResources, string? extraObject) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /XObject << /Bad 9 0 R >> " + extraResources + " >>", maskContent);
        string unsupportedGroup = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f");
        var objects = new List<string> { type3Font, glyphA, outerState, outerMask, unsupportedGroup };
        if (extraObject != null) objects.Add(extraObject);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", objects.ToArray());

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("q /GSZero gs /Bad Do Q 0 0 500 700 re f", "/ExtGState << /GSZero 10 0 R >>", "10 0 obj\n<< /Type /ExtGState /ca 0 >>\nendobj")]
    [InlineData("q 600 600 10 10 re W n /Bad Do Q 0 0 500 700 re f", "", null)]
    public void RenderPage_IgnoresInvisibleImageInsideType3SoftMask(string maskContent, string extraResources, string? extraObject) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /XObject << /Bad 9 0 R >> " + extraResources + " >>", maskContent);
        string unsupportedImage = BuildStreamObject(9, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /Bogus", "x");
        var objects = new List<string> { type3Font, glyphA, outerState, outerMask, unsupportedImage };
        if (extraObject != null) objects.Add(extraObject);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", objects.ToArray());

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_IgnoresRestoredUnsupportedSoftMaskBeforeVisiblePaint() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /ExtGState << /Bad 9 0 R >> >>", "q /Bad gs Q 0 0 500 700 re f");
        string unsupportedState = "9 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 10 0 R >> >>\nendobj";
        string malformedMaskGroup = BuildStreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, unsupportedState, malformedMaskGroup);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("/Bad gs /Sh1 sh", true)]
    [InlineData("0 0 10 10 re W n 20 20 10 10 re W n /Bad gs /Sh1 sh", false)]
    [InlineData("600 600 10 10 re W n /Bad gs /Sh1 sh", false)]
    public void RenderPage_ValidatesNestedSoftMaskWhenDirectShadingPaints(string maskContent, bool expectFallback) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /ExtGState << /Bad 9 0 R >> /Shading << /Sh1 10 0 R >> >>", maskContent);
        string unsupportedState = "9 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 11 0 R >> >>\nendobj";
        string shading = "10 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 1 1] /C1 [1 1 1] /N 1 >> /Extend [true true] >>\nendobj";
        string malformedMaskGroup = BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, unsupportedState, shading, malformedMaskGroup);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Equal(expectFallback, result.CapabilityDiagnostics.Any(diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId));
    }

    [Fact]
    public void RenderPage_IgnoresNestedSoftMaskForFullyTransparentDirectShading() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /ExtGState << /Bad 9 0 R /Zero 12 0 R >> /Shading << /Sh1 10 0 R >> >>", "q /Bad gs /Zero gs /Sh1 sh Q 0 0 500 700 re f");
        string unsupportedState = "9 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 11 0 R >> >>\nendobj";
        string shading = "10 0 obj\n<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 500 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 1 1] /C1 [1 1 1] /N 1 >> /Extend [true true] >>\nendobj";
        string malformedMaskGroup = BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << >>", "0 0 500 700 re f");
        string zeroFillState = "12 0 obj\n<< /Type /ExtGState /ca 0 >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, unsupportedState, shading, malformedMaskGroup, zeroFillState);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_IgnoresStrokeGroupWhenCompositingOpacityIsZero() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /ExtGState << /GSZero 10 0 R >> /XObject << /Bad 9 0 R >> >>", "/GSZero gs /Bad Do");
        string unsupportedGroup = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "10 w 0 0 m 500 700 l S");
        string zeroFillState = "10 0 obj\n<< /Type /ExtGState /ca 0 /CA 1 >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, unsupportedGroup, zeroFillState);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_AppliesSoftMaskGroupMatrixDuringValidation() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [1000 0 1500 700] /Matrix [1 0 0 1 -1000 0] /Group << /S /Transparency /I true >> /Resources << /XObject << /Bad 9 0 R >> >>", "/Bad Do");
        string unsupportedGroup = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [1000 0 1500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "1000 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, unsupportedGroup);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_BoundsSoftMaskType3ValidationAcrossTextBatches() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Font << /Nested 9 0 R >> >>", "BT /Nested 500 Tf (BB) Tj ET");
        string nestedFont = "9 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 10 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << >> >>\nendobj";
        string nestedGlyph = BuildStreamObject(10, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj 60 0 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, nestedFont, nestedGlyph);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxType3GlyphInvocationsPerPage = 2 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.Type3GlyphInvocations, exception.Kind);
    }

    [Fact]
    public void RenderPage_ReusesSoftMaskValidationBudgetAcrossPatternGlyphs() {
        string outerFont = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string outerGlyph = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /Nested 10 0 R >> >>", "BT /Nested 10 Tf (B) Tj ET");
        string nestedFont = "10 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 11 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << /ExtGState << /GS2 12 0 R >> >> >>\nendobj";
        string nestedGlyph = BuildStreamObject(11, "<<", "500 0 d0 /GS2 gs 0 0 500 700 re f");
        string nestedState = "12 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 13 0 R >> >>\nendobj";
        string nestedMask = BuildStreamObject(13, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Font << /Deep 14 0 R >> >>", "BT /Deep 500 Tf (C) Tj ET /Missing Do");
        string deepFont = "14 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /C 15 0 R >> /Encoding << /Differences [67 /C] >> /FirstChar 67 /LastChar 67 /Widths [500] /Resources << >> >>\nendobj";
        string deepGlyph = BuildStreamObject(15, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            outerFont,
            outerGlyph,
            outerState,
            outerMask,
            pattern,
            nestedFont,
            nestedGlyph,
            nestedState,
            nestedMask,
            deepFont,
            deepGlyph);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxType3GlyphInvocationsPerPage = 1 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.Type3GlyphInvocations, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void RenderPage_ReusesSoftMaskTextBudgetAcrossPatternGlyphs() {
        string outerFont = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string outerGlyph = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Font << /Nested 9 0 R >> >>", "BT /Nested 500 Tf (BB) Tj ET /Missing Do");
        string nestedFont = "9 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 10 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << /Pattern << /P1 11 0 R >> >> >>\nendobj";
        string nestedGlyph = BuildStreamObject(10, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(11, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /F1 12 0 R >> >>", "BT /F1 10 Tf (X) Tj ET");
        string helvetica = "12 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            outerFont,
            outerGlyph,
            outerState,
            outerMask,
            nestedFont,
            nestedGlyph,
            pattern,
            helvetica);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedTextCharacters = 1 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.DecodedTextCharacters, exception.Kind);
        Assert.Equal(1, exception.Limit);
        Assert.Equal(2, exception.Actual);
    }

    [Fact]
    public void RenderPage_CarriesSoftMaskDepthIntoStrictPatternCells() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /F1 10 0 R >> >>", "/F1 Do /Missing Do");
        string firstForm = BuildStreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /F2 11 0 R >> >>", "/F2 Do");
        string secondForm = BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >>", "0 0 10 10 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyphA,
            outerState,
            outerMask,
            pattern,
            firstForm,
            secondForm);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 4 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
        Assert.Equal(4, exception.Limit);
    }

    [Fact]
    public void RenderPage_CarriesSoftMaskDepthIntoStrictPatternsNestedInGlyphs() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Font << /Nested 9 0 R >> >>", "BT /Nested 500 Tf (B) Tj ET");
        string nestedFont = "9 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 10 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << /Pattern << /P1 11 0 R >> >> >>\nendobj";
        string nestedGlyph = BuildStreamObject(10, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(11, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /F1 12 0 R >> >>", "/F1 Do");
        string nestedForm = BuildStreamObject(12, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >>", "0 0 10 10 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyphA,
            outerState,
            outerMask,
            nestedFont,
            nestedGlyph,
            pattern,
            nestedForm);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 4 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
        Assert.Equal(4, exception.Limit);
    }

    [Fact]
    public void RenderPage_DoesNotReusePatternMaskDepthAcrossNestedCalls() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R /None 8 0 R >> /XObject << /Nested 9 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f /None gs /Nested Do");
        string maskedState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 10 0 R >> >>\nendobj";
        string clearState = "8 0 obj\n<< /Type /ExtGState /SMask /None >>\nendobj";
        string nestedCall = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << /ExtGState << /GS1 7 0 R >> >>", "/GS1 gs 0 0 500 700 re f");
        string maskGroup = BuildStreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 11 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(11, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << /F1 12 0 R >> >>", "/F1 Do");
        string nestedForm = BuildStreamObject(12, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << >>", "0 0 10 10 re f");
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyphA,
            maskedState,
            clearState,
            nestedCall,
            maskGroup,
            pattern,
            nestedForm);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxContentNestingDepth = 4 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.Pages[0].ToDrawing());

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
        Assert.Equal(4, exception.Limit);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnsupportedTilingPatternPayloadInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /XObject << >> >>", "/Missing Do");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForNestedSoftMaskPayloadInsideStrictPatternCell() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /ExtGState << /NestedMask 10 0 R >> >>", "/NestedMask gs 0 0 10 10 re f");
        string nestedState = "10 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 11 0 R >> >>\nendobj";
        string nestedMask = BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Group << /S /Transparency /I true >> /Resources << /XObject << >> >>", "/Missing Do");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, pattern, nestedState, nestedMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForUnsupportedOrdinaryTextInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "BT /Missing 48 Tf 0 100 Td (A) Tj ET");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenLuminositySoftMaskDependsOnInheritedColor() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 1 0 0 rg /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForMissingClippingTextFontInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "BT /Missing 48 Tf 7 Tr 0 100 Td (A) Tj ET 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_ProjectsSupportedOrdinaryTextInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> >>", "BT /FText 300 Tf 25 200 Td (A) Tj ET");
        string ordinaryFont = "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, ordinaryFont);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_UsesSeparateTextBudgetForSoftMaskValidation() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 1 0 0 rg 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> >>", "BT /FText 300 Tf 25 200 Td (A) Tj ET");
        string ordinaryFont = "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf(
            "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            "<< /Font << /FType3 5 0 R >> >>",
            type3Font,
            glyphA,
            outerState,
            outerMask,
            ordinaryFont);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            // The live output contains the page Type 3 character and one mask character.
            // Validation receives its own budget and must not charge either character twice.
            Limits = new PdfReadLimits { MaxDecodedTextCharacters = 2 }
        });

        OfficeDrawing drawing = document.Pages[0].ToDrawing();

        Assert.Contains(drawing.Elements, element => element is OfficeDrawingEffectGroup { SoftMask: not null });
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    public void RenderPage_FailsClosedForUnsupportedOrdinaryTextPaintModeInsideType3SoftMask(int renderingMode) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> >>", $"BT /FText 300 Tf {renderingMode} Tr 25 200 Td (A) Tj ET");
        string ordinaryFont = "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, ordinaryFont);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_DoesNotDoubleChargeColoredPatternType3GlyphValidation() {
        string outerFont = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string outerGlyph = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Font << /Nested 10 0 R >> >>", "BT /Nested 10 Tf (B) Tj ET");
        string nestedFont = "10 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 11 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << >> >>\nendobj";
        string nestedGlyph = BuildStreamObject(11, "<<", "500 0 d0 0 0 500 700 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", outerFont, outerGlyph, outerState, outerMask, pattern, nestedFont, nestedGlyph);
        var options = new PdfReadOptions {
            Limits = new PdfReadLimits { MaxType3GlyphInvocationsPerPage = 2 }
        };

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf, readOptions: options));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Theory]
    [InlineData("Symbol")]
    [InlineData("ZapfDingbats")]
    public void RenderPage_FailsClosedForUnembeddedSymbolicFontInsideType3SoftMask(string baseFont) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> >>", "BT /FText 300 Tf 25 200 Td (A) Tj ET");
        string ordinaryFont = $"9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /{baseFont} >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, ordinaryFont);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData("HelveticaNeue")]
    [InlineData("Times-New-Roman")]
    [InlineData("CourierPrime")]
    public void RenderPage_FailsClosedForUnembeddedNonBase14FontInsideType3SoftMask(string baseFont) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> >>", "BT /FText 300 Tf 25 200 Td (A) Tj ET");
        string ordinaryFont = $"9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /{baseFont} /Encoding /WinAnsiEncoding >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, ordinaryFont);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_IgnoresNonpaintingUnsupportedTextInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> >>", "1 1 1 rg 0 0 500 700 re f BT /FText 48 Tf 3 Tr 25 200 Td (A) Tj ET");
        string ordinaryFont = "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /HelveticaNeue /Encoding /WinAnsiEncoding >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, ordinaryFont);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_IgnoresTransparentUnsupportedTextInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> /ExtGState << /Zero 10 0 R >> >>", "1 1 1 rg 0 0 500 700 re f /Zero gs BT /FText 48 Tf 25 200 Td (A) Tj ET");
        string unsupportedFont = "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /HelveticaNeue /Encoding /WinAnsiEncoding >>\nendobj";
        string transparentFill = "10 0 obj\n<< /Type /ExtGState /ca 0 >>\nendobj";
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, unsupportedFont, transparentFill);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_IgnoresActiveEffectForNonpaintingType3TextInsideSoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /ExtGState << /Bad 9 0 R >> /Font << /Nested 11 0 R >> >>", "q /Bad gs BT /Nested 500 Tf (B) Tj ET Q 0 0 500 700 re f");
        string unsupportedState = "9 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 10 0 R >> >>\nendobj";
        string malformedMaskGroup = BuildStreamObject(10, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << >>", "0 0 500 700 re f");
        string nestedFont = "11 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 12 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << >> >>\nendobj";
        string emptyGlyph = BuildStreamObject(12, "<<", "500 0 d0");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, unsupportedState, malformedMaskGroup, nestedFont, emptyGlyph);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForPatternPaintedImageMaskInsideColoredType3SoftMaskText() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Font << /Nested 9 0 R >> >>", "BT /Nested 500 Tf (B) Tj ET");
        string nestedFont = "9 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 10 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << /Pattern << /P1 11 0 R >> /XObject << /Im1 12 0 R >> >> >>\nendobj";
        string nestedGlyph = BuildStreamObject(10, "<<", "500 0 d0 /Pattern cs /P1 scn q 500 0 0 700 0 0 cm /Im1 Do Q");
        string pattern = BuildStreamObject(11, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f");
        string imageMask = BuildStreamObject(12, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, nestedFont, nestedGlyph, pattern, imageMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_FailsClosedForPatternPaintedImageMaskInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> /XObject << /Im1 10 0 R >> >>", "/Pattern cs /P1 scn q 500 0 0 700 0 0 cm /Im1 Do Q");
        string pattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 0 0 rg 0 0 5 5 re f");
        string imageMask = BuildStreamObject(10, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, pattern, imageMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_InheritsInvokingFormResourcesForNestedSoftMaskGroup() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /XObject << /Nested 9 0 R >> >>", "/Nested Do");
        string nestedForm = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Resources << /ExtGState << /NestedMask 10 0 R >> /XObject << /Im1 12 0 R >> >>", "/NestedMask gs 0 0 500 700 re f");
        string nestedState = "10 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 11 0 R >> >>\nendobj";
        string nestedMask = BuildStreamObject(11, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >>", "q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(12, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", "\u00ff");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, nestedForm, nestedState, nestedMask, image);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.NotEqual(OfficeColor.Transparent, raster.GetPixel(24, 94));
    }

    [Fact]
    public void RenderPage_FailsClosedForUnresolvedImageTransparencyInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /XObject << /Im1 9 0 R >> >>", "q 500 0 0 700 0 0 cm /Im1 Do Q");
        string image = BuildStreamObject(9, "<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Mask 10 0 R", "abcdef");
        string explicitMask = BuildStreamObject(10, "<< /Type /XObject /Subtype /Image /Width 2 /Height 1 /ImageMask true /BitsPerComponent 1", "@");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, image, explicitMask);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Fact]
    public void RenderPage_ProjectsRecursivelyNestedPatternsInsideType3SoftMaskText() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 1 0 0 rg 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Font << /Nested 9 0 R >> >>", "BT /Nested 500 Tf (B) Tj ET");
        string nestedFont = "9 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /B 10 0 R >> /Encoding << /Differences [66 /B] >> /FirstChar 66 /LastChar 66 /Widths [500] /Resources << /Pattern << /P1 11 0 R >> >> >>\nendobj";
        string nestedGlyph = BuildStreamObject(10, "<<", "500 0 d0 /Pattern cs /P1 scn 0 0 500 700 re f");
        string outerPattern = BuildStreamObject(11, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Pattern << /P2 12 0 R >> >>", "/Pattern cs /P2 scn 0 0 10 10 re f");
        string innerPattern = BuildStreamObject(12, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 g 0 0 5 5 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, nestedFont, nestedGlyph, outerPattern, innerPattern);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingSoftMask softMask = Assert.Single(
            drawing.Elements.OfType<OfficeDrawingEffectGroup>(),
            group => group.SoftMask != null).SoftMask!;

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        int patternCount = CountTilingPatterns(softMask.Drawing);
        Assert.True(patternCount >= 2, "Expected two nested pattern drawings, found " + patternCount + ".");
    }

    [Fact]
    public void RenderPage_ProjectsRecursivelyNestedPatternsDirectlyInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 1 0 0 rg 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> >>", "/Pattern cs /P1 scn 0 0 500 700 re f");
        string outerPattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << /Pattern << /P2 10 0 R >> >>", "/Pattern cs /P2 scn 0 0 10 10 re f");
        string innerPattern = BuildStreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 g 0 0 5 5 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, outerPattern, innerPattern);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeDrawingSoftMask softMask = Assert.Single(
            drawing.Elements.OfType<OfficeDrawingEffectGroup>(),
            group => group.SoftMask != null).SoftMask!;

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.True(CountTilingPatterns(softMask.Drawing) >= 2);
    }

    [Fact]
    public void RenderPage_FailsClosedForInvalidDirectPatternSelectionInsideType3SoftMask() {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true >> /Resources << /Pattern << /P1 9 0 R >> >>", "/Pattern cs 1 /P1 scn 0 0 500 700 re f");
        string pattern = BuildStreamObject(9, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 g 0 0 5 5 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("1")]
    [InlineData("2")]
    public void RenderPage_FailsClosedForPatternPaintedOrdinaryTextInsideType3SoftMask(string textRenderingMode) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Luminosity /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << /Font << /FText 9 0 R >> /Pattern << /P1 10 0 R >> >>", $"/Pattern cs /P1 scn /Pattern CS /P1 SCN BT /FText 300 Tf {textRenderingMode} Tr 25 200 Td (A) Tj ET");
        string ordinaryFont = "9 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>\nendobj";
        string pattern = BuildStreamObject(10, "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>", "1 g 0 0 5 5 re f");
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", type3Font, glyphA, outerState, outerMask, ordinaryFont, pattern);

        AssertType3FallsBackWithoutNativeShapes(pdf);
    }

    private static int CountTilingPatterns(OfficeDrawing drawing) {
        int count = 0;
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingTilingPattern pattern) {
                count += 1 + CountTilingPatterns(pattern.Tile);
            } else if (element is OfficeDrawingEffectGroup group) {
                count += CountTilingPatterns(group.Drawing);
                if (group.SoftMask != null) count += CountTilingPatterns(group.SoftMask.Drawing);
            } else if (element is OfficeDrawingGroup clippedGroup) {
                count += CountTilingPatterns(clippedGroup.Drawing);
            }
        }
        return count;
    }
}
