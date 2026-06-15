using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlParagraphFrameTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Paragraph_Frame_Metadata() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Framed").SetFrame(frame => {
            frame.SetSize(widthTwips: 3600, heightTwips: 0)
                .SetAnchors(RtfParagraphFrameHorizontalAnchor.Margin, RtfParagraphFrameVerticalAnchor.Paragraph)
                .SetPosition(RtfParagraphFrameHorizontalPosition.NegativeAbsolute, horizontalTwips: -180, RtfParagraphFrameVerticalPosition.Absolute, verticalTwips: 720)
                .SetWrapping(noWrap: true, allDirectionsTwips: 120, horizontalTwips: 240, verticalTwips: 360, overlayText: true, noOverlap: true)
                .SetDropCap(2, RtfDropCapKind.InText);
            frame.AnchorLocked = true;
        });
        document.AddParagraph("Plain");

        string html = document.ToHtml(new RtfToHtmlOptions { NewLine = "\n" });

        Assert.Contains("data-officeimo-rtf-paragraph-frame=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadRtfFromHtml();
        Assert.Equal(2, roundTrip.Paragraphs.Count);
        RtfParagraph framed = roundTrip.Paragraphs[0];
        AssertFrame(
            framed.Frame,
            widthTwips: 3600,
            heightTwips: 0,
            horizontalAnchor: RtfParagraphFrameHorizontalAnchor.Margin,
            verticalAnchor: RtfParagraphFrameVerticalAnchor.Paragraph,
            horizontalPosition: RtfParagraphFrameHorizontalPosition.NegativeAbsolute,
            horizontalTwips: -180,
            verticalPosition: RtfParagraphFrameVerticalPosition.Absolute,
            verticalTwips: 720,
            anchorLocked: true,
            noOverlap: true,
            noWrap: true,
            wrapDistance: 120,
            wrapDistanceHorizontal: 240,
            wrapDistanceVertical: 360,
            overlayText: true,
            dropCapLines: 2,
            dropCapKind: RtfDropCapKind.InText);
        Assert.False(roundTrip.Paragraphs[1].Frame.HasAnyValue);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\absw3600\absh0\phmrg\posnegx-180\pvpara\posy720\abslock\absnoovrlp1\nowrap\dxfrtext120\dfrmtxtx240\dfrmtxty360\overlay\dropcapli2\dropcapt1", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Stylesheet_Frame_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfStyle style = document.AddStyle(12, "Framed Block");
        style.SetFrame(frame => {
            frame.SetSize(widthTwips: 5040, heightTwips: -720)
                .SetAnchors(RtfParagraphFrameHorizontalAnchor.Page, RtfParagraphFrameVerticalAnchor.Page)
                .SetPosition(RtfParagraphFrameHorizontalPosition.Center, null, RtfParagraphFrameVerticalPosition.Top, null)
                .SetWrapping(noWrap: true, allDirectionsTwips: 173, horizontalTwips: 240, verticalTwips: 360, overlayText: true, noOverlap: false)
                .SetDropCap(3, RtfDropCapKind.Margin);
            frame.AnchorLocked = true;
        });
        document.AddParagraph("Styled").SetStyle(style.Id);

        string html = document.ToHtml(new RtfToHtmlOptions { FragmentOnly = false, NewLine = "\n" });

        Assert.Contains("<meta name=\"officeimo-rtf-styles\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadRtfFromHtml();
        RtfStyle roundTripStyle = Assert.Single(roundTrip.Styles);
        Assert.Equal(style.Id, roundTripStyle.Id);
        AssertFrame(
            roundTripStyle.Frame,
            widthTwips: 5040,
            heightTwips: -720,
            horizontalAnchor: RtfParagraphFrameHorizontalAnchor.Page,
            verticalAnchor: RtfParagraphFrameVerticalAnchor.Page,
            horizontalPosition: RtfParagraphFrameHorizontalPosition.Center,
            horizontalTwips: null,
            verticalPosition: RtfParagraphFrameVerticalPosition.Top,
            verticalTwips: null,
            anchorLocked: true,
            noOverlap: false,
            noWrap: true,
            wrapDistance: 173,
            wrapDistanceHorizontal: 240,
            wrapDistanceVertical: 360,
            overlayText: true,
            dropCapLines: 3,
            dropCapKind: RtfDropCapKind.Margin);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\absw5040\absh-720\phpg\posxc\pvpg\posyt\abslock\absnoovrlp0\nowrap\dxfrtext173\dfrmtxtx240\dfrmtxty360\overlay\dropcapli3\dropcapt2", rtf, StringComparison.Ordinal);
    }

    private static void AssertFrame(
        RtfParagraphFrame frame,
        int widthTwips,
        int heightTwips,
        RtfParagraphFrameHorizontalAnchor horizontalAnchor,
        RtfParagraphFrameVerticalAnchor verticalAnchor,
        RtfParagraphFrameHorizontalPosition horizontalPosition,
        int? horizontalTwips,
        RtfParagraphFrameVerticalPosition verticalPosition,
        int? verticalTwips,
        bool anchorLocked,
        bool? noOverlap,
        bool noWrap,
        int wrapDistance,
        int wrapDistanceHorizontal,
        int wrapDistanceVertical,
        bool overlayText,
        int dropCapLines,
        RtfDropCapKind dropCapKind) {
        Assert.True(frame.HasAnyValue);
        Assert.Equal(widthTwips, frame.WidthTwips);
        Assert.Equal(heightTwips, frame.HeightTwips);
        Assert.Equal(horizontalAnchor, frame.HorizontalAnchor);
        Assert.Equal(verticalAnchor, frame.VerticalAnchor);
        Assert.Equal(horizontalPosition, frame.HorizontalPosition);
        Assert.Equal(horizontalTwips, frame.HorizontalPositionTwips);
        Assert.Equal(verticalPosition, frame.VerticalPosition);
        Assert.Equal(verticalTwips, frame.VerticalPositionTwips);
        Assert.Equal(anchorLocked, frame.AnchorLocked);
        Assert.Equal(noOverlap, frame.NoOverlap);
        Assert.Equal(noWrap, frame.NoWrap);
        Assert.Equal(wrapDistance, frame.TextWrapDistanceTwips);
        Assert.Equal(wrapDistanceHorizontal, frame.TextWrapDistanceHorizontalTwips);
        Assert.Equal(wrapDistanceVertical, frame.TextWrapDistanceVerticalTwips);
        Assert.Equal(overlayText, frame.OverlayText);
        Assert.Equal(dropCapLines, frame.DropCapLines);
        Assert.Equal(dropCapKind, frame.DropCapKind);
    }
}
