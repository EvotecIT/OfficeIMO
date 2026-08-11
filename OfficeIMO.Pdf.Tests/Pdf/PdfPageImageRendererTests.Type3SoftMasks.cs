using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("/Pattern cs /Missing scn 0 0 500 700 re f")]
    [InlineData("/Pattern CS /Missing SCN 10 w 0 0 m 500 700 l S")]
    public void RenderPage_FailsClosedForMissingDirectPatternInsideType3SoftMask(string maskContent) {
        string type3Font = "5 0 obj\n<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /A 6 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [500] /Resources << /ExtGState << /GS1 7 0 R >> >> >>\nendobj";
        string glyphA = BuildStreamObject(6, "<<", "500 0 d0 /GS1 gs 0 0 500 700 re f");
        string outerState = "7 0 obj\n<< /Type /ExtGState /SMask << /S /Alpha /G 8 0 R >> >>\nendobj";
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency >> /Resources << /Pattern << >> >>", maskContent);
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
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency >> /Resources << >>", maskContent);
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
        string outerMask = BuildStreamObject(8, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency >> /Resources << /XObject << /Bad 9 0 R >> " + extraResources + " >>", maskContent);
        string unsupportedGroup = BuildStreamObject(9, "<< /Type /XObject /Subtype /Form /BBox [0 0 500 700] /Group << /S /Transparency /I true /CS /DeviceRGB >> /Resources << >>", "0 0 500 700 re f");
        var objects = new List<string> { type3Font, glyphA, outerState, outerMask, unsupportedGroup };
        if (extraObject != null) objects.Add(extraObject);
        byte[] pdf = BuildSingleStreamPdf("BT /FType3 18 Tf 20 100 Td (A) Tj ET", "<< /Font << /FType3 5 0 R >> >>", objects.ToArray());

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }
}
