using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlBookmarkTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Bookmark_Anchors_In_Inline_Order() {
        const string html = "<p><a id=\"Anchor\" data-officeimo-rtf-bookmark=\"start\"></a>Bookmarked<a data-officeimo-rtf-bookmark=\"end\" data-officeimo-rtf-bookmark-name=\"Anchor\"></a> text</p>";

        RtfDocument document = html.ToRtfDocument();
        RtfParagraph paragraph = Assert.Single(document.Paragraphs);

        Assert.Collection(paragraph.Inlines,
            inline => {
                RtfBookmarkMarker marker = Assert.IsType<RtfBookmarkMarker>(inline);
                Assert.Equal(RtfBookmarkMarkerKind.Start, marker.Kind);
                Assert.Equal("Anchor", marker.Name);
            },
            inline => Assert.Equal("Bookmarked", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfBookmarkMarker marker = Assert.IsType<RtfBookmarkMarker>(inline);
                Assert.Equal(RtfBookmarkMarkerKind.End, marker.Kind);
                Assert.Equal("Anchor", marker.Name);
            },
            inline => Assert.Equal(" text", Assert.IsType<RtfRun>(inline).Text));

        string rtf = document.ToRtf();
        Assert.Contains(@"{\*\bkmkstart Anchor}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\bkmkend Anchor}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Legacy_Id_Anchor_As_Bookmark_Start() {
        const string html = "<p><a id=\"Target\"></a>Target text</p>";

        RtfParagraph paragraph = Assert.Single(html.ToRtfDocument().Paragraphs);

        RtfBookmarkMarker marker = Assert.IsType<RtfBookmarkMarker>(paragraph.Inlines[0]);
        Assert.Equal(RtfBookmarkMarkerKind.Start, marker.Kind);
        Assert.Equal("Target", marker.Name);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Bookmark_Metadata() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddBookmarkStart("Anchor");
        paragraph.AddText("Bookmarked");
        paragraph.AddBookmarkEnd("Anchor");
        paragraph.AddText(" text");

        string html = document.ToHtml();

        Assert.Equal("<p><a id=\"Anchor\" data-officeimo-rtf-bookmark=\"start\" data-officeimo-rtf-bookmark-name=\"Anchor\"></a>Bookmarked<a data-officeimo-rtf-bookmark=\"end\" data-officeimo-rtf-bookmark-name=\"Anchor\"></a> text</p>", html);

        RtfParagraph roundTripParagraph = Assert.Single(html.ToRtfDocument().Paragraphs);
        Assert.Collection(roundTripParagraph.Inlines,
            inline => Assert.Equal(RtfBookmarkMarkerKind.Start, Assert.IsType<RtfBookmarkMarker>(inline).Kind),
            inline => Assert.Equal("Bookmarked", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBookmarkMarkerKind.End, Assert.IsType<RtfBookmarkMarker>(inline).Kind),
            inline => Assert.Equal(" text", Assert.IsType<RtfRun>(inline).Text));
    }
}
