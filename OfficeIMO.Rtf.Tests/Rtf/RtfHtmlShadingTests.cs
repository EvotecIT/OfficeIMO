using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlShadingTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Paragraph_And_Run_Shading_Metadata() {
        const string html = "<p style=\"background-color:#e6f2ff;--officeimo-rtf-shading-foreground:#4472c4;--officeimo-rtf-shading-percent:6250;--officeimo-rtf-shading-pattern:bgdkfdiag\">Assessment <span style=\"background-color:#fff2cc;--officeimo-rtf-shading-foreground:#00aa55;--officeimo-rtf-shading-percent:37.5%;--officeimo-rtf-shading-pattern:dark-diagonal-cross\">flag</span></p>";

        RtfDocument document = html.ToRtfDocument();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        AssertColor(document, paragraph.BackgroundColorIndex, 0xE6, 0xF2, 0xFF);
        AssertColor(document, paragraph.ShadingForegroundColorIndex, 0x44, 0x72, 0xC4);
        Assert.Equal(6250, paragraph.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, paragraph.ShadingPattern);
        RtfRun flag = Assert.Single(paragraph.Runs, run => run.Text == "flag");
        AssertColor(document, flag.CharacterBackgroundColorIndex, 0xFF, 0xF2, 0xCC);
        AssertColor(document, flag.CharacterShadingForegroundColorIndex, 0x00, 0xAA, 0x55);
        Assert.Equal(3750, flag.CharacterShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkDiagonalCross, flag.CharacterShadingPattern);

        string rtf = document.ToRtf();
        Assert.Contains(@"\cfpat", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\shading6250", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\bgdkfdiag", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chcfpat", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chshdng3750", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chbgdkdcross", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripParagraph = Assert.Single(RtfDocument.Read(rtf).Document.Paragraphs);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, roundTripParagraph.ShadingPattern);
        Assert.Equal(6250, roundTripParagraph.ShadingPatternPercent);
        RtfRun roundTripFlag = Assert.Single(roundTripParagraph.Runs, run => run.Text == "flag");
        Assert.Equal(RtfShadingPattern.DarkDiagonalCross, roundTripFlag.CharacterShadingPattern);
        Assert.Equal(3750, roundTripFlag.CharacterShadingPatternPercent);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Paragraph_And_Run_Shading_Metadata() {
        RtfDocument document = RtfDocument.Create();
        int paragraphBackground = document.AddColor(0xE6, 0xF2, 0xFF);
        int paragraphForeground = document.AddColor(0x44, 0x72, 0xC4);
        int runBackground = document.AddColor(0xFF, 0xF2, 0xCC);
        int runForeground = document.AddColor(0x00, 0xAA, 0x55);
        RtfParagraph paragraph = document.AddParagraph()
            .SetShading(paragraphBackground, paragraphForeground, patternPercent: 6250, pattern: RtfShadingPattern.DarkForwardDiagonal);
        paragraph.AddText("Assessment ");
        paragraph.AddText("flag")
            .SetCharacterShading(runBackground, runForeground, patternPercent: 3750, pattern: RtfShadingPattern.DarkDiagonalCross);

        string html = document.ToHtml(RtfToHtmlOptions.CreateRoundTripProfile());

        Assert.Equal("<p style=\"background-color:#E6F2FF;--officeimo-rtf-shading-foreground:#4472C4;--officeimo-rtf-shading-percent:6250;--officeimo-rtf-shading-pattern:dark-forward-diagonal;\">Assessment <span style=\"background-color:#FFF2CC;--officeimo-rtf-shading-foreground:#00AA55;--officeimo-rtf-shading-percent:3750;--officeimo-rtf-shading-pattern:dark-diagonal-cross;\">flag</span></p>", html);

        RtfParagraph roundTripParagraph = Assert.Single(html.ToRtfDocument().Paragraphs);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, roundTripParagraph.ShadingPattern);
        Assert.Equal(6250, roundTripParagraph.ShadingPatternPercent);
        RtfRun roundTripFlag = Assert.Single(roundTripParagraph.Runs, run => run.Text == "flag");
        Assert.Equal(RtfShadingPattern.DarkDiagonalCross, roundTripFlag.CharacterShadingPattern);
        Assert.Equal(3750, roundTripFlag.CharacterShadingPatternPercent);
    }

    private static void AssertColor(RtfDocument document, int? colorIndex, byte red, byte green, byte blue) {
        Assert.True(colorIndex.HasValue);
        RtfColor color = document.Colors[colorIndex.Value - 1];
        Assert.Equal(red, color.Red);
        Assert.Equal(green, color.Green);
        Assert.Equal(blue, color.Blue);
    }
}
