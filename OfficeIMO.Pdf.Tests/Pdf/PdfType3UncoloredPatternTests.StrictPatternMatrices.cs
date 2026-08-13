using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
    [Fact]
    public void RenderPage_FailsClosedWhenStrictPatternMatrixHasNoFiniteInverse() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Matrix [1e308 0 0 1e308 1e308 0] /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenPatternStrokeConsumesInheritedLineState() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "20 w [3 1] 0 d 1 J 2 j /Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "0 0 5 5 re S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_AllowsPatternStrokeWithCompleteLocalLineState() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 w [] 0 d 0 J 0 j 0 0 5 5 re S");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenNestedType3StrokeConsumesPatternLineState() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "20 w [3 1] 0 d 1 J 2 j /Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << /Font << /Nested 10 0 R >> >>",
            patternContent: "BT /Nested 5 Tf 0 0 Td (A) Tj ET",
            extraObjects: new[] {
                "10 0 obj\n<< /Type /Font /Subtype /Type3 /PaintType 1 /FontBBox [0 0 5 5] /FontMatrix [1 0 0 1 0 0] /CharProcs << /A 11 0 R >> /Encoding << /Differences [65 /A] >> /FirstChar 65 /LastChar 65 /Widths [5] /Resources << >> >>\nendobj",
                StreamObject(11, "<<", "0 0 5 5 re S")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForSurplusStrictXObjectOperands() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "99 /Missing Do");

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }
}
