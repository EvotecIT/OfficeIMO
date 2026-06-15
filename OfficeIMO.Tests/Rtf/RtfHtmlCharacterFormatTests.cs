using System;
using System.Linq;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlCharacterFormatTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Character_Border() {
        const string html = "<p><span style=\"border:1pt solid #0c2238\">Flag</span><span style=\"border-top:2pt dashed red\"> Side</span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun flagged = paragraph.Runs.Single(run => run.Text == "Flag");
        Assert.Equal(RtfParagraphBorderStyle.Single, flagged.CharacterBorder.Style);
        Assert.Equal(20, flagged.CharacterBorder.Width);
        Assert.Equal(1, flagged.CharacterBorder.ColorIndex);

        RtfRun sideOnly = paragraph.Runs.Single(run => run.Text == " Side");
        Assert.False(sideOnly.CharacterBorder.HasAnyValue);

        string rtf = document.ToRtf();
        Assert.Contains(@"\chbrdr\brdrs\brdrw20\brdrcf1", rtf, StringComparison.Ordinal);

        RtfRun roundTripFlagged = RtfDocument.Read(rtf).Document.Paragraphs[0].Runs.Single(run => run.Text == "Flag");
        Assert.Equal(RtfParagraphBorderStyle.Single, roundTripFlagged.CharacterBorder.Style);
        Assert.Equal(20, roundTripFlagged.CharacterBorder.Width);
        Assert.Equal(1, roundTripFlagged.CharacterBorder.ColorIndex);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Character_Border() {
        RtfDocument document = RtfDocument.Create();
        int dark = document.AddColor(12, 34, 56);
        document.AddParagraph().AddText("Flag")
            .SetCharacterBorder(RtfParagraphBorderStyle.Double, width: 40, colorIndex: dark);

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"border:2pt double #0C2238;\">Flag</span></p>", html);

        RtfRun roundTripFlagged = html.ToRtfDocumentFromHtml().Paragraphs[0].Runs.Single(run => run.Text == "Flag");
        Assert.Equal(RtfParagraphBorderStyle.Double, roundTripFlagged.CharacterBorder.Style);
        Assert.Equal(40, roundTripFlagged.CharacterBorder.Width);
        Assert.Equal(1, roundTripFlagged.CharacterBorder.ColorIndex);
    }
}
