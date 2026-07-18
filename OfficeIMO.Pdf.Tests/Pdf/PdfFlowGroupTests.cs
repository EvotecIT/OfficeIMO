using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfFlowGroupTests {
    [Fact]
    public void KeepTogether_MovesMeasuredGroupAndCapturesFinalRegion() {
        var capture = new PdfLayoutPositionCapture();
        byte[] bytes = CreateShortPageDocument()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .Spacer(105)
            .Flow(
                item => item.Rectangle(60, 80),
                new PdfFlowOptions { KeepTogether = true },
                capture)
            .ToBytes();

        PdfLayoutRegion region = Assert.Single(capture.Regions);
        Assert.Equal(2, PdfInspector.Inspect(bytes).PageCount);
        Assert.Equal(2, region.PageNumber);
        Assert.Equal(80D, region.Height, 3);
        Assert.False(capture.WasSkipped);
    }

    [Fact]
    public void Deferred_ReplaysAfterMinimumSpaceStartsANewPage() {
        var capture = new PdfLayoutPositionCapture();
        int invocations = 0;
        byte[] bytes = CreateShortPageDocument()
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .Spacer(100)
            .Deferred(
                context => {
                    invocations++;
                    return item => item.Paragraph(paragraph => paragraph.Text("Deferred page " + context.PageNumber));
                },
                new PdfFlowOptions { MinimumRemainingHeight = 100 },
                capture)
            .ToBytes();

        Assert.Equal(2, invocations);
        Assert.Contains("Deferred page 2", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        Assert.Equal(2, capture.Last!.PageNumber);
    }

    [Fact]
    public void ConditionalSkipAndStopPoliciesControlNestedAndFollowingFlow() {
        var conditionalCapture = new PdfLayoutPositionCapture();
        bool materialized = false;
        byte[] conditional = CreateShortPageDocument()
            .Deferred(
                _ => {
                    materialized = true;
                    return item => item.Paragraph(paragraph => paragraph.Text("Hidden"));
                },
                new PdfFlowOptions { ShowIf = _ => false },
                conditionalCapture)
            .Paragraph(paragraph => paragraph.Text("Visible"))
            .ToBytes();

        var skipCapture = new PdfLayoutPositionCapture();
        byte[] skipped = CreateShortPageDocument()
            .Paragraph(paragraph => paragraph.Text("Before"))
            .Spacer(105)
            .Flow(
                item => item.Rectangle(60, 80),
                new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.Skip },
                skipCapture)
            .Paragraph(paragraph => paragraph.Text("After skip"))
            .ToBytes();

        byte[] stopped = CreateShortPageDocument()
            .Paragraph(paragraph => paragraph.Text("Before"))
            .Spacer(105)
            .Flow(
                item => item.Rectangle(60, 80),
                new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.StopDocument })
            .Paragraph(paragraph => paragraph.Text("After stop"))
            .ToBytes();

        Assert.False(materialized);
        Assert.True(conditionalCapture.WasSkipped);
        Assert.Contains("Visible", PdfReadDocument.Open(conditional).ExtractText(), StringComparison.Ordinal);
        Assert.True(skipCapture.WasSkipped);
        Assert.Contains("After skip", PdfReadDocument.Open(skipped).ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain("After stop", PdfReadDocument.Open(stopped).ExtractText(), StringComparison.Ordinal);
    }

    private static PdfDocument CreateShortPageDocument() {
        return PdfDocument.Create(new PdfOptions {
            PageWidth = 200,
            PageHeight = 200,
            MarginLeft = 20,
            MarginTop = 20,
            MarginRight = 20,
            MarginBottom = 20,
            CompressContentStreams = false
        });
    }
}
