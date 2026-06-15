using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlParagraphFormatTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraph_Borders() {
        const string html = "<p style=\"border:1pt solid #0c2238;border-left:2pt double red\">Boxed</p><p>Plain</p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfParagraph boxed = document.Paragraphs[0];
        Assert.Equal(RtfParagraphBorderStyle.Single, boxed.TopBorder.Style);
        Assert.Equal(20, boxed.TopBorder.Width);
        Assert.Equal(1, boxed.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, boxed.LeftBorder.Style);
        Assert.Equal(40, boxed.LeftBorder.Width);
        Assert.Equal(2, boxed.LeftBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Single, boxed.RightBorder.Style);
        Assert.False(document.Paragraphs[1].TopBorder.HasAnyValue);

        string rtf = document.ToRtf();
        Assert.Contains(@"\brdrt\brdrs\brdrw20\brdrcf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\brdrl\brdrdb\brdrw40\brdrcf2", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripBoxed = RtfDocument.Read(rtf).Document.Paragraphs[0];
        Assert.Equal(RtfParagraphBorderStyle.Single, roundTripBoxed.TopBorder.Style);
        Assert.Equal(RtfParagraphBorderStyle.Double, roundTripBoxed.LeftBorder.Style);
        Assert.Equal(40, roundTripBoxed.LeftBorder.Width);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Paragraph_Borders() {
        RtfDocument document = RtfDocument.Create();
        int dark = document.AddColor(12, 34, 56);
        int red = document.AddColor(255, 0, 0);
        document.AddParagraph("Boxed")
            .SetBorder(RtfParagraphBorderSide.Top, RtfParagraphBorderStyle.Single, width: 20, colorIndex: dark)
            .SetBorder(RtfParagraphBorderSide.Left, RtfParagraphBorderStyle.Double, width: 40, colorIndex: red)
            .SetBorder(RtfParagraphBorderSide.Bottom, RtfParagraphBorderStyle.Dotted);

        string html = document.ToHtml();

        Assert.Equal("<p style=\"border-top:1pt solid #0C2238;border-left:2pt double #FF0000;border-bottom:dotted;\">Boxed</p>", html);

        RtfParagraph roundTripBoxed = html.ToRtfDocumentFromHtml().Paragraphs[0];
        Assert.Equal(RtfParagraphBorderStyle.Single, roundTripBoxed.TopBorder.Style);
        Assert.Equal(20, roundTripBoxed.TopBorder.Width);
        Assert.Equal(1, roundTripBoxed.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, roundTripBoxed.LeftBorder.Style);
        Assert.Equal(40, roundTripBoxed.LeftBorder.Width);
        Assert.Equal(2, roundTripBoxed.LeftBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Dotted, roundTripBoxed.BottomBorder.Style);
    }
}
