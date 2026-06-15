using System;
using System.Linq;
using OfficeIMO.Rtf;
using OfficeIMO.Html.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlCharacterFormatTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Character_Border() {
        const string html = "<p><span style=\"border:1pt solid #0c2238\">Flag</span><span style=\"border-top:2pt dashed red\"> Side</span></p>";

        RtfDocument document = html.LoadFromHtml();

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

        RtfRun roundTripFlagged = html.LoadFromHtml().Paragraphs[0].Runs.Single(run => run.Text == "Flag");
        Assert.Equal(RtfParagraphBorderStyle.Double, roundTripFlagged.CharacterBorder.Style);
        Assert.Equal(40, roundTripFlagged.CharacterBorder.Width);
        Assert.Equal(1, roundTripFlagged.CharacterBorder.ColorIndex);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Rich_Underline_Style_And_Color() {
        const string html = "<p><span style=\"text-decoration-line:underline;text-decoration-style:wavy;text-decoration-color:#0c2238\">Wave</span><span style=\"text-decoration-style:double;text-decoration-color:red\"> Plain</span></p>";

        RtfDocument document = html.LoadFromHtml();

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

        RtfRun roundTripFlagged = html.LoadFromHtml().Paragraphs[0].Runs.Single(run => run.Text == "Flag");
        Assert.Equal(RtfUnderlineStyle.ThickDashDotDot, roundTripFlagged.UnderlineStyle);
        Assert.Equal(1, roundTripFlagged.UnderlineColorIndex);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Double_Strike() {
        const string html = "<p><span style=\"text-decoration-line:line-through;text-decoration-style:double\">Double</span><span style=\"text-decoration-style:double\"> Plain</span></p>";

        RtfDocument document = html.LoadFromHtml();

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

        RtfRun roundTripDoubled = html.LoadFromHtml().Paragraphs[0].Runs.Single(run => run.Text == "Double");
        Assert.True(roundTripDoubled.DoubleStrike);
        Assert.False(roundTripDoubled.Strike);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Caps_And_SmallCaps() {
        const string html = "<p><span style=\"text-transform:uppercase\">Caps</span><span style=\"font-variant-caps:small-caps\"> Small</span><span style=\"text-transform:uppercase\"><span style=\"text-transform:none\"> Plain</span></span></p>";

        RtfDocument document = html.LoadFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.Equal(RtfCapsStyle.Caps, paragraph.Runs.Single(run => run.Text == "Caps").CapsStyle);
        Assert.Equal(RtfCapsStyle.SmallCaps, paragraph.Runs.Single(run => run.Text == " Small").CapsStyle);
        Assert.Equal(RtfCapsStyle.None, paragraph.Runs.Single(run => run.Text == " Plain").CapsStyle);

        string rtf = document.ToRtf();
        Assert.Contains(@"\caps Caps", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\scaps  Small", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripParagraph = RtfDocument.Read(rtf).Document.Paragraphs[0];
        Assert.Equal(RtfCapsStyle.Caps, roundTripParagraph.Runs.Single(run => run.Text == "Caps").CapsStyle);
        Assert.Equal(RtfCapsStyle.SmallCaps, roundTripParagraph.Runs.Single(run => run.Text == " Small").CapsStyle);
        Assert.Equal(RtfCapsStyle.None, roundTripParagraph.Runs.Single(run => run.Text == " Plain").CapsStyle);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Caps_And_SmallCaps() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Caps").SetCapsStyle(RtfCapsStyle.Caps);
        paragraph.AddText(" Small").SetCapsStyle(RtfCapsStyle.SmallCaps);

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"text-transform:uppercase;--officeimo-rtf-caps-style:caps;\">Caps</span><span style=\"font-variant-caps:small-caps;--officeimo-rtf-caps-style:small-caps;\"> Small</span></p>", html);

        RtfParagraph roundTripParagraph = html.LoadFromHtml().Paragraphs[0];
        Assert.Equal(RtfCapsStyle.Caps, roundTripParagraph.Runs.Single(run => run.Text == "Caps").CapsStyle);
        Assert.Equal(RtfCapsStyle.SmallCaps, roundTripParagraph.Runs.Single(run => run.Text == " Small").CapsStyle);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Character_Metrics() {
        const string html = "<p><span style=\"letter-spacing:2pt;font-stretch:80%;vertical-align:3pt\">Raised<span style=\"letter-spacing:normal;font-stretch:100%;vertical-align:baseline\"> Plain</span></span><span style=\"letter-spacing:-1pt\"> Condensed</span></p>";

        RtfDocument document = html.LoadFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun raised = paragraph.Runs.Single(run => run.Text == "Raised");
        Assert.Equal(40, raised.CharacterSpacingTwips);
        Assert.Equal(80, raised.CharacterScalePercent);
        Assert.Equal(6, raised.CharacterOffsetHalfPoints);

        RtfRun plain = paragraph.Runs.Single(run => run.Text == " Plain");
        Assert.Null(plain.CharacterSpacingTwips);
        Assert.Null(plain.CharacterScalePercent);
        Assert.Null(plain.CharacterOffsetHalfPoints);

        RtfRun condensed = paragraph.Runs.Single(run => run.Text == " Condensed");
        Assert.Equal(-20, condensed.CharacterSpacingTwips);

        string rtf = document.ToRtf();
        Assert.Contains(@"\expndtw40", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\charscalex80", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\up6", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\expndtw-20", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripParagraph = RtfDocument.Read(rtf).Document.Paragraphs[0];
        Assert.Equal(40, roundTripParagraph.Runs.Single(run => run.Text == "Raised").CharacterSpacingTwips);
        Assert.Equal(80, roundTripParagraph.Runs.Single(run => run.Text == "Raised").CharacterScalePercent);
        Assert.Equal(6, roundTripParagraph.Runs.Single(run => run.Text == "Raised").CharacterOffsetHalfPoints);
        Assert.Null(roundTripParagraph.Runs.Single(run => run.Text == " Plain").CharacterSpacingTwips);
        Assert.Equal(-20, roundTripParagraph.Runs.Single(run => run.Text == " Condensed").CharacterSpacingTwips);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Character_Metrics() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Raised")
            .SetCharacterSpacingTwips(40)
            .SetCharacterScale(80)
            .SetCharacterOffsetHalfPoints(6);
        paragraph.AddText(" Lowered").SetCharacterOffsetHalfPoints(-4);

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"letter-spacing:2pt;font-stretch:80%;--officeimo-rtf-character-scale:80;vertical-align:3pt;--officeimo-rtf-character-offset:6;\">Raised</span><span style=\"vertical-align:-2pt;--officeimo-rtf-character-offset:-4;\"> Lowered</span></p>", html);

        RtfParagraph roundTripParagraph = html.LoadFromHtml().Paragraphs[0];
        RtfRun raised = roundTripParagraph.Runs.Single(run => run.Text == "Raised");
        Assert.Equal(40, raised.CharacterSpacingTwips);
        Assert.Equal(80, raised.CharacterScalePercent);
        Assert.Equal(6, raised.CharacterOffsetHalfPoints);
        Assert.Equal(-4, roundTripParagraph.Runs.Single(run => run.Text == " Lowered").CharacterOffsetHalfPoints);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Character_Effects() {
        const string html = "<p><span style=\"visibility:hidden\">Hidden</span><span style=\"--officeimo-rtf-outline:true\"> Outline</span><span style=\"text-shadow:1pt 1pt 0 currentColor\"> Shadow</span><span style=\"--officeimo-rtf-emboss:true\"> Emboss</span><span style=\"--officeimo-rtf-imprint:true\"> Imprint</span><span style=\"visibility:hidden\"><span style=\"visibility:visible\"> Plain</span></span></p>";

        RtfDocument document = html.LoadFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        Assert.True(paragraph.Runs.Single(run => run.Text == "Hidden").Hidden);
        Assert.True(paragraph.Runs.Single(run => run.Text == " Outline").Outline);
        Assert.True(paragraph.Runs.Single(run => run.Text == " Shadow").Shadow);
        Assert.True(paragraph.Runs.Single(run => run.Text == " Emboss").Emboss);
        Assert.True(paragraph.Runs.Single(run => run.Text == " Imprint").Imprint);
        Assert.False(paragraph.Runs.Single(run => run.Text == " Plain").Hidden);

        string rtf = document.ToRtf();
        Assert.Contains(@"\v Hidden\v0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\outl  Outline\outl0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\shad  Shadow\shad0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\embo  Emboss\embo0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\impr  Imprint\impr0", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripParagraph = RtfDocument.Read(rtf).Document.Paragraphs[0];
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == "Hidden").Hidden);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Outline").Outline);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Shadow").Shadow);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Emboss").Emboss);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Imprint").Imprint);
        Assert.False(roundTripParagraph.Runs.Single(run => run.Text == " Plain").Hidden);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Character_Effects() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Hidden").SetHidden();
        paragraph.AddText(" Outline").SetOutline();
        paragraph.AddText(" Shadow").SetShadow();
        paragraph.AddText(" Emboss").SetEmboss();
        paragraph.AddText(" Imprint").SetImprint();

        string html = document.ToHtml();

        Assert.Equal("<p><span style=\"visibility:hidden;--officeimo-rtf-hidden:true;\">Hidden</span><span style=\"--officeimo-rtf-outline:true;\"> Outline</span><span style=\"text-shadow:1pt 1pt 0 currentColor;--officeimo-rtf-shadow:true;\"> Shadow</span><span style=\"--officeimo-rtf-emboss:true;\"> Emboss</span><span style=\"--officeimo-rtf-imprint:true;\"> Imprint</span></p>", html);

        RtfParagraph roundTripParagraph = html.LoadFromHtml().Paragraphs[0];
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == "Hidden").Hidden);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Outline").Outline);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Shadow").Shadow);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Emboss").Emboss);
        Assert.True(roundTripParagraph.Runs.Single(run => run.Text == " Imprint").Imprint);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Language_And_Direction() {
        const string html = "<p><span dir=\"rtl\" lang=\"ar-SA\">RTL</span><span style=\"direction:ltr;--officeimo-rtf-lang:1045\"> Polish</span><span dir=\"auto\"> Plain</span></p>";

        RtfDocument document = html.LoadFromHtml();

        RtfParagraph paragraph = Assert.Single(document.Paragraphs);
        RtfRun rtl = paragraph.Runs.Single(run => run.Text == "RTL");
        Assert.Equal(RtfTextDirection.RightToLeft, rtl.Direction);
        Assert.Equal(1025, rtl.LanguageId);

        RtfRun polish = paragraph.Runs.Single(run => run.Text == " Polish");
        Assert.Equal(RtfTextDirection.LeftToRight, polish.Direction);
        Assert.Equal(1045, polish.LanguageId);

        RtfRun plain = paragraph.Runs.Single(run => run.Text == " Plain");
        Assert.Null(plain.Direction);
        Assert.Null(plain.LanguageId);

        string rtf = document.ToRtf();
        Assert.Contains(@"\rtlch \lang1025 RTL", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ltrch \lang1045  Polish", rtf, StringComparison.Ordinal);

        RtfParagraph roundTripParagraph = RtfDocument.Read(rtf).Document.Paragraphs[0];
        Assert.Equal(RtfTextDirection.RightToLeft, roundTripParagraph.Runs.Single(run => run.Text == "RTL").Direction);
        Assert.Equal(1025, roundTripParagraph.Runs.Single(run => run.Text == "RTL").LanguageId);
        Assert.Equal(RtfTextDirection.LeftToRight, roundTripParagraph.Runs.Single(run => run.Text == " Polish").Direction);
        Assert.Equal(1045, roundTripParagraph.Runs.Single(run => run.Text == " Polish").LanguageId);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Language_And_Direction() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("RTL")
            .SetDirection(RtfTextDirection.RightToLeft)
            .SetLanguage(1025);
        paragraph.AddText(" Polish")
            .SetDirection(RtfTextDirection.LeftToRight)
            .SetLanguage(1045);

        string html = document.ToHtml();

        Assert.Equal("<p><span lang=\"ar-SA\" dir=\"rtl\" style=\"--officeimo-rtf-lang:1025;direction:rtl;unicode-bidi:isolate;--officeimo-rtf-direction:rtl;\">RTL</span><span lang=\"pl-PL\" dir=\"ltr\" style=\"--officeimo-rtf-lang:1045;direction:ltr;unicode-bidi:isolate;--officeimo-rtf-direction:ltr;\"> Polish</span></p>", html);

        RtfParagraph roundTripParagraph = html.LoadFromHtml().Paragraphs[0];
        Assert.Equal(RtfTextDirection.RightToLeft, roundTripParagraph.Runs.Single(run => run.Text == "RTL").Direction);
        Assert.Equal(1025, roundTripParagraph.Runs.Single(run => run.Text == "RTL").LanguageId);
        Assert.Equal(RtfTextDirection.LeftToRight, roundTripParagraph.Runs.Single(run => run.Text == " Polish").Direction);
        Assert.Equal(1045, roundTripParagraph.Runs.Single(run => run.Text == " Polish").LanguageId);
    }
}
