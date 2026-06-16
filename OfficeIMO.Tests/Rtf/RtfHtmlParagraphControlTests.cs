using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlParagraphControlTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Paragraph_Controls_And_Tab_Stops() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Patient\tResult\t12.34");
        paragraph
            .SetPagination(keepWithNext: true, keepLinesTogether: true, widowControl: false, suppressLineNumbers: true, autoHyphenation: false)
            .SetContextualSpacing(false)
            .SetAdjustRightIndent(true)
            .SetSnapToLineGrid(false)
            .SetParagraphSpacing(beforeAuto: true, afterAuto: false);
        paragraph.AddTabStop(1440);
        paragraph.AddTabStop(2880, RtfTabAlignment.Right, RtfTabLeader.Dots);
        paragraph.AddTabStop(4320, RtfTabAlignment.Decimal, RtfTabLeader.MiddleDots);

        string html = document.ToHtml(new RtfToHtmlOptions { NewLine = "\n" });

        Assert.Contains("data-officeimo-rtf-paragraph-controls=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocument();
        RtfParagraph roundTripParagraph = Assert.Single(roundTrip.Paragraphs);
        Assert.True(roundTripParagraph.KeepWithNext);
        Assert.True(roundTripParagraph.KeepLinesTogether);
        Assert.True(roundTripParagraph.SuppressLineNumbers);
        Assert.Equal(false, roundTripParagraph.AutoHyphenation);
        Assert.Equal(false, roundTripParagraph.ContextualSpacing);
        Assert.Equal(true, roundTripParagraph.AdjustRightIndent);
        Assert.Equal(false, roundTripParagraph.SnapToLineGrid);
        Assert.Equal(false, roundTripParagraph.WidowControl);
        Assert.Equal(true, roundTripParagraph.SpaceBeforeAuto);
        Assert.Equal(false, roundTripParagraph.SpaceAfterAuto);
        Assert.Collection(roundTripParagraph.TabStops,
            tabStop => {
                Assert.Equal(1440, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Left, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.None, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(2880, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Right, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.Dots, tabStop.Leader);
            },
            tabStop => {
                Assert.Equal(4320, tabStop.PositionTwips);
                Assert.Equal(RtfTabAlignment.Decimal, tabStop.Alignment);
                Assert.Equal(RtfTabLeader.MiddleDots, tabStop.Leader);
            });

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\keepn", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\keep", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\noline", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\hyphpar0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\contextualspace0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\adjustright", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nosnaplinegrid", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\nowidctlpar", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sbauto1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\saauto0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tx1440", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tldot\tqr\tx2880", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tlmdot\tqdec\tx4320", rtf, StringComparison.Ordinal);
    }
}
