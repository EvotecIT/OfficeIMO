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

    [Fact]
    public void Html_ToRtfDocument_Parses_Rich_Underline_Style_And_Color() {
        const string html = "<p><span style=\"text-decoration-line:underline;text-decoration-style:wavy;text-decoration-color:#0c2238\">Wave</span><span style=\"text-decoration-style:double;text-decoration-color:red\"> Plain</span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun wave = paragraph.Runs.Single(run => run.Text == "Wave");
        Assert.Equal(RtfUnderlineStyle.Wave, wave.UnderlineStyle);
        Assert.Equal(1, wave.UnderlineColorIndex);

        RtfRun plain = paragraph.Runs.Single(run => run.Text == " Plain");
        Assert.Equal(RtfUnderlineStyle.None, plain.UnderlineStyle);
        Assert.Null(plain.UnderlineColorIndex);
        Assert.False(plain.DoubleStrike);

        string rtf = document.ToRtf();
        Assert.Contains(@"\ulwave", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ulc1", rtf, StringComparison.Ordinal);

        RtfRun roundTripWave = RtfDocument.Read(rtf).Document.Paragraphs[0].Runs.Single(run => run.Text == "Wave");
        Assert.Equal(RtfUnderlineStyle.Wave, roundTripWave.UnderlineStyle);
        Assert.Equal(1, roundTripWave.UnderlineColorIndex);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Rich_Underline_Style_And_Color() {
        RtfDocument document = RtfDocument.Create();
        int dark = document.AddColor(12, 34, 56);
        document.AddParagraph().AddText("Flag")
            .SetUnderline(RtfUnderlineStyle.ThickDashDotDot)
            .SetUnderlineColor(dark);

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"text-decoration-line:underline;text-decoration-style:dashed;--officeimo-rtf-underline-style:thick-dash-dot-dot;text-decoration-color:#0C2238;\">Flag</span></p>", html);

        RtfRun roundTripFlagged = html.ToRtfDocumentFromHtml().Paragraphs[0].Runs.Single(run => run.Text == "Flag");
        Assert.Equal(RtfUnderlineStyle.ThickDashDotDot, roundTripFlagged.UnderlineStyle);
        Assert.Equal(1, roundTripFlagged.UnderlineColorIndex);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Double_Strike() {
        const string html = "<p><span style=\"text-decoration-line:line-through;text-decoration-style:double\">Double</span><span style=\"text-decoration-style:double\"> Plain</span></p>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun doubled = paragraph.Runs.Single(run => run.Text == "Double");
        Assert.True(doubled.DoubleStrike);
        Assert.False(doubled.Strike);

        RtfRun plain = paragraph.Runs.Single(run => run.Text == " Plain");
        Assert.False(plain.Strike);
        Assert.False(plain.DoubleStrike);

        string rtf = document.ToRtf();
        Assert.Contains(@"\striked Double\striked0", rtf, StringComparison.Ordinal);

        RtfRun roundTripDoubled = RtfDocument.Read(rtf).Document.Paragraphs[0].Runs.Single(run => run.Text == "Double");
        Assert.True(roundTripDoubled.DoubleStrike);
        Assert.False(roundTripDoubled.Strike);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Double_Strike() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddText("Double").SetDoubleStrike();

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"text-decoration-line:line-through;text-decoration-style:double;--officeimo-rtf-strike-style:double;\">Double</span></p>", html);

        RtfRun roundTripDoubled = html.ToRtfDocumentFromHtml().Paragraphs[0].Runs.Single(run => run.Text == "Double");
        Assert.True(roundTripDoubled.DoubleStrike);
        Assert.False(roundTripDoubled.Strike);
    }
}
