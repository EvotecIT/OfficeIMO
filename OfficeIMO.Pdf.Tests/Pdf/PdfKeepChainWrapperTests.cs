using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfKeepChainWrapperTests {
    // Page body is 130pt: a 70pt filler plus three 20pt lines fill 130pt exactly.
    [Theory]
    [InlineData("plain")]
    [InlineData("semantic")]
    [InlineData("flow")]
    [InlineData("nested")]
    public void HeadingAtWrapperEndKeepsWithTheBlockAfterTheWrapper(string wrapper) {
        byte[] bytes = PdfDocument.Create(builder => builder.Content(flow => {
            flow.AddParagraph(p => p.AddText("FillerMarker"), style: Line(70, keepWithNext: false));
            flow.AddParagraph(p => p.AddText("PrecedingMarker"), style: Line(20, keepWithNext: true));
            Action<PdfContentBuilder> children = nested => {
                nested.AddParagraph(p => p.AddText("BoundaryMarker"), style: Line(20, keepWithNext: false));
                nested.H2("HeadingMarker", style: new PdfHeadingStyle {
                    FontSize = P(12), LineSpacing = PdfLineSpacing.Exact(P(20)),
                    SpacingBefore = P(0), SpacingAfter = P(0), KeepWithNext = true
                });
            };
            if (wrapper == "semantic") flow.Semantic(PdfSemanticRole.Section, children);
            else if (wrapper == "flow") flow.Flow(children);
            else if (wrapper == "nested") flow.Semantic(PdfSemanticRole.Section, content => content.Flow(children));
            else children(flow);
            flow.AddParagraph(p => p.AddText("FollowingMarker"), style: Line(20, keepWithNext: false));
        }), Options()).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        // The chain ends at the non-keeping boundary paragraph inside the wrapper.
        Assert.Contains("PrecedingMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("BoundaryMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("HeadingMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("HeadingMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("FollowingMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Equal(2, Assert.Single(PdfInspector.Inspect(bytes).Outlines).PageNumber);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void KeepWithNextReservesAKeptFlowAsAWhole(bool row) {
        Action<PdfContentBuilder> children = flow => {
            flow.AddParagraph(p => p.AddText("FillerMarker"), style: Line(80, keepWithNext: false));
            flow.AddParagraph(p => p.AddText("PrecedingMarker"), style: Line(20, keepWithNext: true));
            flow.Flow(group => {
                group.AddParagraph(p => p.AddText("GroupFirstMarker"), style: Line(20, keepWithNext: false));
                group.AddParagraph(p => p.AddText("GroupLastMarker"), style: Line(20, keepWithNext: false));
            }, new PdfFlowOptions { KeepTogether = true });
        };
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => {
            if (row) content.Row(columns => columns.RelativeColumn(children));
            else children(content);
        }), Options()).ToBytes();

        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("PrecedingMarker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("PrecedingMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("GroupFirstMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("GroupLastMarker", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    private static global::OfficeIMO.OfficeLength P(double points) => global::OfficeIMO.OfficeLength.Points(points);

    private static PdfParagraphStyle Line(double height, bool keepWithNext) => new() {
        FontSize = P(12), LineSpacing = PdfLineSpacing.Exact(P(height)),
        SpacingBefore = P(0), SpacingAfter = P(0), KeepWithNext = keepWithNext, KeepTogether = true, WidowControl = false
    };

    private static PdfCreateOptions Options() => new() {
        PageSetup = {
            Size = new PdfPageSize(P(400), P(180)),
            Margins = { Left = P(25), Right = P(25), Top = P(25), Bottom = P(25) }
        },
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = P(12), CreateOutlineFromHeadings = true
    };
}
