using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfCharacterFormattingTests {
    [Theory]
    [InlineData(@"\ul", RtfUnderlineStyle.Single)]
    [InlineData(@"\ulw", RtfUnderlineStyle.Words)]
    [InlineData(@"\uldb", RtfUnderlineStyle.Double)]
    [InlineData(@"\uld", RtfUnderlineStyle.Dotted)]
    [InlineData(@"\uldash", RtfUnderlineStyle.Dash)]
    [InlineData(@"\uldashd", RtfUnderlineStyle.DashDot)]
    [InlineData(@"\uldashdd", RtfUnderlineStyle.DashDotDot)]
    [InlineData(@"\ulth", RtfUnderlineStyle.Thick)]
    [InlineData(@"\ulthd", RtfUnderlineStyle.ThickDotted)]
    [InlineData(@"\ulthdash", RtfUnderlineStyle.ThickDash)]
    [InlineData(@"\ulthdashd", RtfUnderlineStyle.ThickDashDot)]
    [InlineData(@"\ulthdashdd", RtfUnderlineStyle.ThickDashDotDot)]
    [InlineData(@"\ulwave", RtfUnderlineStyle.Wave)]
    [InlineData(@"\ulhwave", RtfUnderlineStyle.HeavyWave)]
    [InlineData(@"\uldbwave", RtfUnderlineStyle.DoubleWave)]
    [InlineData(@"\ulldash", RtfUnderlineStyle.LongDash)]
    [InlineData(@"\ulthldash", RtfUnderlineStyle.ThickLongDash)]
    public void Read_Binds_Rich_Underline_Styles(string control, RtfUnderlineStyle expectedStyle) {
        string rtf = $@"{{\rtf1\ansi\pard {control} Text\ulnone plain\par}}";

        RtfReadResult result = RtfDocument.Read(rtf);

        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Contains(paragraph.Runs, run => run.Text == "Text" && run.UnderlineStyle == expectedStyle);
        RtfRun plain = paragraph.Runs.Single(run => run.Text == "plain");
        Assert.Equal(RtfUnderlineStyle.None, plain.UnderlineStyle);
        Assert.False(plain.Underline);
    }

    [Theory]
    [InlineData(@"\ul", RtfUnderlineStyle.Single)]
    [InlineData(@"\ulw", RtfUnderlineStyle.Words)]
    [InlineData(@"\uldb", RtfUnderlineStyle.Double)]
    [InlineData(@"\uld", RtfUnderlineStyle.Dotted)]
    [InlineData(@"\uldash", RtfUnderlineStyle.Dash)]
    [InlineData(@"\uldashd", RtfUnderlineStyle.DashDot)]
    [InlineData(@"\uldashdd", RtfUnderlineStyle.DashDotDot)]
    [InlineData(@"\ulth", RtfUnderlineStyle.Thick)]
    [InlineData(@"\ulthd", RtfUnderlineStyle.ThickDotted)]
    [InlineData(@"\ulthdash", RtfUnderlineStyle.ThickDash)]
    [InlineData(@"\ulthdashd", RtfUnderlineStyle.ThickDashDot)]
    [InlineData(@"\ulthdashdd", RtfUnderlineStyle.ThickDashDotDot)]
    [InlineData(@"\ulwave", RtfUnderlineStyle.Wave)]
    [InlineData(@"\ulhwave", RtfUnderlineStyle.HeavyWave)]
    [InlineData(@"\uldbwave", RtfUnderlineStyle.DoubleWave)]
    [InlineData(@"\ulldash", RtfUnderlineStyle.LongDash)]
    [InlineData(@"\ulthldash", RtfUnderlineStyle.ThickLongDash)]
    public void Write_Emits_Rich_Underline_Styles(string expectedControl, RtfUnderlineStyle style) {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph().AddText("Text").SetUnderline(style);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(expectedControl + " Text", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ulnone", rtf, StringComparison.Ordinal);
        RtfRun run = Assert.Single(Assert.Single(result.Document.Paragraphs).Runs, item => item.Text == "Text");
        Assert.Equal(style, run.UnderlineStyle);
    }

    [Fact]
    public void Write_And_Read_Underline_Color_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Styled").SetUnderline(RtfUnderlineStyle.Double).SetUnderlineColor(red);
        paragraph.AddText(" Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\uldb", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ulc1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ulnone", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ulc0", rtf, StringComparison.Ordinal);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        RtfRun styled = readParagraph.Runs.Single(run => run.Text == "Styled");
        Assert.Equal(RtfUnderlineStyle.Double, styled.UnderlineStyle);
        Assert.Equal(red, styled.UnderlineColorIndex);
        RtfRun plain = readParagraph.Runs.Single(run => run.Text == " Plain");
        Assert.Equal(RtfUnderlineStyle.None, plain.UnderlineStyle);
        Assert.Null(plain.UnderlineColorIndex);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Rich_Underline_Style_And_Color() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        WordParagraph styled = paragraph.AddText("Styled").SetUnderline(UnderlineValues.WavyDouble);
        styled._run!.RunProperties!.Underline!.Color = "4472C4";
        paragraph.AddText("Plain");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfRun rtfRun = Assert.Single(Assert.Single(rtfDocument.Paragraphs).Runs, run => run.Text == "Styled");
        Assert.Equal(RtfUnderlineStyle.DoubleWave, rtfRun.UnderlineStyle);
        Assert.Equal(1, rtfRun.UnderlineColorIndex);
        Assert.Contains(@"\uldbwave", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ulc1", rtf, StringComparison.Ordinal);
        WordParagraph roundTripRun = Assert.Single(roundTrip.Paragraphs, run => run.Text == "Styled");
        Assert.Equal(UnderlineValues.WavyDouble, roundTripRun.Underline);
        Assert.Equal("4472C4", roundTripRun._run?.RunProperties?.Underline?.Color?.Value);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Rich_Underline_Style_And_Color() {
        RtfDocument rtfDocument = RtfDocument.Create();
        int colorId = rtfDocument.AddColor(0x44, 0x72, 0xC4);
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        paragraph.AddText("Styled").SetUnderline(RtfUnderlineStyle.ThickDashDotDot).SetUnderlineColor(colorId);
        paragraph.AddText("Plain");

        using WordDocument word = rtfDocument.ToWordDocument();

        WordParagraph styled = Assert.Single(word.Paragraphs, run => run.Text == "Styled");
        Assert.Equal(UnderlineValues.DashDotDotHeavy, styled.Underline);
        Assert.Equal("4472C4", styled._run?.RunProperties?.Underline?.Color?.Value);
        WordParagraph plain = Assert.Single(word.Paragraphs, run => run.Text == "Plain");
        Assert.Null(plain.Underline);
        Assert.Null(plain._run?.RunProperties?.Underline?.Color?.Value);
    }

    [Fact]
    public void Read_Binds_Character_Shading_And_Border_Without_Leaking() {
        const string rtf = @"{\rtf1\ansi{\colortbl;\red230\green242\blue255;\red68\green114\blue196;}\pard \chcbpat1\chcfpat2\chshdng3750\chbgdkfdiag\chbrdr\brdrdb\brdrw12\brdrcf2 Styled\plain Plain\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        RtfRun styled = paragraph.Runs.Single(run => run.Text == "Styled");
        Assert.Equal(1, styled.CharacterBackgroundColorIndex);
        Assert.Equal(2, styled.CharacterShadingForegroundColorIndex);
        Assert.Equal(3750, styled.CharacterShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, styled.CharacterShadingPattern);
        Assert.Equal(RtfParagraphBorderStyle.Double, styled.CharacterBorder.Style);
        Assert.Equal(12, styled.CharacterBorder.Width);
        Assert.Equal(2, styled.CharacterBorder.ColorIndex);

        RtfRun plain = paragraph.Runs.Single(run => run.Text == "Plain");
        Assert.Null(plain.CharacterBackgroundColorIndex);
        Assert.Null(plain.CharacterShadingForegroundColorIndex);
        Assert.Null(plain.CharacterShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.None, plain.CharacterShadingPattern);
        Assert.False(plain.CharacterBorder.HasAnyValue);
    }

    [Fact]
    public void Write_And_Read_Character_Shading_And_Border_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        int fill = document.AddColor(0xE6, 0xF2, 0xFF);
        int foreground = document.AddColor(0x44, 0x72, 0xC4);
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Styled")
            .SetCharacterShading(fill, foregroundColorIndex: foreground, patternPercent: 6250, pattern: RtfShadingPattern.DarkDiagonalCross)
            .SetCharacterBorder(RtfParagraphBorderStyle.Dashed, width: 14, colorIndex: foreground);
        paragraph.AddText(" Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\chcbpat1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chcfpat2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chshdng6250", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chbgdkdcross", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\chbrdr\brdrdash\brdrw14\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\plain  Plain", rtf, StringComparison.Ordinal);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        RtfRun styled = readParagraph.Runs.Single(run => run.Text == "Styled");
        Assert.Equal(fill, styled.CharacterBackgroundColorIndex);
        Assert.Equal(foreground, styled.CharacterShadingForegroundColorIndex);
        Assert.Equal(6250, styled.CharacterShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkDiagonalCross, styled.CharacterShadingPattern);
        Assert.Equal(RtfParagraphBorderStyle.Dashed, styled.CharacterBorder.Style);
        Assert.Equal(14, styled.CharacterBorder.Width);
        Assert.Equal(foreground, styled.CharacterBorder.ColorIndex);

        RtfRun plain = readParagraph.Runs.Single(run => run.Text == " Plain");
        Assert.Null(plain.CharacterBackgroundColorIndex);
        Assert.Null(plain.CharacterShadingForegroundColorIndex);
        Assert.Null(plain.CharacterShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.None, plain.CharacterShadingPattern);
        Assert.False(plain.CharacterBorder.HasAnyValue);
    }

    [Fact]
    public void Read_Binds_Character_Spacing_Scale_Kerning_And_Offset() {
        const string rtf = @"{\rtf1\ansi\pard \expndtw40\charscalex80\kerning24\up6 Raised\expndtw0\charscalex100\kerning0\up0 plain \expnd-4 Condensed\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        RtfRun raised = paragraph.Runs.Single(run => run.Text == "Raised");
        Assert.Equal(40, raised.CharacterSpacingTwips);
        Assert.Equal(80, raised.CharacterScalePercent);
        Assert.Equal(24, raised.KerningHalfPoints);
        Assert.Equal(6, raised.CharacterOffsetHalfPoints);

        RtfRun plain = paragraph.Runs.Single(run => run.Text == "plain ");
        Assert.Null(plain.CharacterSpacingTwips);
        Assert.Null(plain.CharacterScalePercent);
        Assert.Null(plain.KerningHalfPoints);
        Assert.Null(plain.CharacterOffsetHalfPoints);

        RtfRun condensed = paragraph.Runs.Single(run => run.Text == "Condensed");
        Assert.Equal(-20, condensed.CharacterSpacingTwips);
    }

    [Fact]
    public void Write_And_Read_Character_Spacing_Scale_Kerning_And_Offset_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Raised")
            .SetCharacterSpacingTwips(40)
            .SetCharacterScale(80)
            .SetKerningHalfPoints(24)
            .SetCharacterOffsetHalfPoints(6);
        paragraph.AddText(" Lowered").SetCharacterOffsetHalfPoints(-4);
        paragraph.AddText(" plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\expndtw40", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\charscalex80", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\kerning24", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\up6", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\dn4", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\expndtw0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\charscalex100", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\kerning0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\up0", rtf, StringComparison.Ordinal);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        RtfRun raised = readParagraph.Runs.Single(run => run.Text == "Raised");
        Assert.Equal(40, raised.CharacterSpacingTwips);
        Assert.Equal(80, raised.CharacterScalePercent);
        Assert.Equal(24, raised.KerningHalfPoints);
        Assert.Equal(6, raised.CharacterOffsetHalfPoints);
        Assert.Equal(-4, readParagraph.Runs.Single(run => run.Text == " Lowered").CharacterOffsetHalfPoints);
        RtfRun plain = readParagraph.Runs.Single(run => run.Text == " plain");
        Assert.Null(plain.CharacterSpacingTwips);
        Assert.Null(plain.CharacterScalePercent);
        Assert.Null(plain.KerningHalfPoints);
        Assert.Null(plain.CharacterOffsetHalfPoints);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Character_Spacing_Scale_Kerning_And_Offset() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        WordParagraph styled = paragraph.AddText("Styled").SetSpacing(40);
        styled._run!.RunProperties ??= new RunProperties();
        styled._run.RunProperties.CharacterScale = new CharacterScale { Val = 80 };
        styled._run.RunProperties.Kern = new Kern { Val = 24U };
        styled._run.RunProperties.Position = new Position { Val = "6" };
        paragraph.AddText("Plain");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfRun rtfRun = Assert.Single(Assert.Single(rtfDocument.Paragraphs).Runs, run => run.Text == "Styled");
        Assert.Equal(40, rtfRun.CharacterSpacingTwips);
        Assert.Equal(80, rtfRun.CharacterScalePercent);
        Assert.Equal(24, rtfRun.KerningHalfPoints);
        Assert.Equal(6, rtfRun.CharacterOffsetHalfPoints);
        Assert.Contains(@"\expndtw40", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\charscalex80", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\kerning24", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\up6", rtf, StringComparison.Ordinal);

        WordParagraph roundTripRun = Assert.Single(roundTrip.Paragraphs, run => run.Text == "Styled");
        Assert.Equal(40, roundTripRun.Spacing);
        Assert.Equal(80, roundTripRun._run?.RunProperties?.CharacterScale?.Val?.Value);
        Assert.Equal(24U, roundTripRun._run?.RunProperties?.Kern?.Val?.Value);
        Assert.Equal("6", roundTripRun._run?.RunProperties?.Position?.Val?.Value);
        WordParagraph plain = Assert.Single(roundTrip.Paragraphs, run => run.Text == "Plain");
        Assert.Null(plain.Spacing);
        Assert.Null(plain._run?.RunProperties?.CharacterScale);
        Assert.Null(plain._run?.RunProperties?.Kern);
        Assert.Null(plain._run?.RunProperties?.Position);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Character_Spacing_Scale_Kerning_And_Offset() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        paragraph.AddText("Styled")
            .SetCharacterSpacingTwips(-30)
            .SetCharacterScale(120)
            .SetKerningHalfPoints(18)
            .SetCharacterOffsetHalfPoints(-5);
        paragraph.AddText("Plain");

        using WordDocument word = rtfDocument.ToWordDocument();

        WordParagraph styled = Assert.Single(word.Paragraphs, run => run.Text == "Styled");
        Assert.Equal(-30, styled.Spacing);
        Assert.Equal(120, styled._run?.RunProperties?.CharacterScale?.Val?.Value);
        Assert.Equal(18U, styled._run?.RunProperties?.Kern?.Val?.Value);
        Assert.Equal("-5", styled._run?.RunProperties?.Position?.Val?.Value);
        WordParagraph plain = Assert.Single(word.Paragraphs, run => run.Text == "Plain");
        Assert.Null(plain.Spacing);
        Assert.Null(plain._run?.RunProperties?.CharacterScale);
        Assert.Null(plain._run?.RunProperties?.Kern);
        Assert.Null(plain._run?.RunProperties?.Position);
    }

    [Fact]
    public void Read_Binds_Super_And_Sub_Zero_As_Baseline_Reset() {
        const string rtf = @"{\rtf1\ansi\pard \super Raised\super0 Plain \sub Lowered\sub0 Base\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal(RtfVerticalPosition.Superscript, paragraph.Runs.Single(run => run.Text == "Raised").VerticalPosition);
        Assert.Equal(RtfVerticalPosition.Baseline, paragraph.Runs.Single(run => run.Text == "Plain ").VerticalPosition);
        Assert.Equal(RtfVerticalPosition.Subscript, paragraph.Runs.Single(run => run.Text == "Lowered").VerticalPosition);
        Assert.Equal(RtfVerticalPosition.Baseline, paragraph.Runs.Single(run => run.Text == "Base").VerticalPosition);
    }

    [Fact]
    public void Write_And_Read_Default_And_Run_Language_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultLanguage(1045);
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Default ");
        paragraph.AddText("English").SetLanguage(1033);
        paragraph.AddText(" Polish");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.StartsWith(@"{\rtf1\ansi\deff0\deflang1045", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\lang1033 English", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\lang1045  Polish", rtf, StringComparison.Ordinal);
        Assert.Equal(1045, result.Document.Settings.DefaultLanguageId);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal(1045, readParagraph.Runs.Single(run => run.Text == "Default ").LanguageId);
        Assert.Equal(1033, readParagraph.Runs.Single(run => run.Text == "English").LanguageId);
        Assert.Equal(1045, readParagraph.Runs.Single(run => run.Text == " Polish").LanguageId);
    }

    [Fact]
    public void Read_Binds_Default_And_Run_Language() {
        const string rtf = @"{\rtf1\ansi\deff0\deflang1045\pard Default \lang1033 English\lang1045 Polish\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.Equal(1045, result.Document.Settings.DefaultLanguageId);
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal(1045, paragraph.Runs.Single(run => run.Text == "Default ").LanguageId);
        Assert.Equal(1033, paragraph.Runs.Single(run => run.Text == "English").LanguageId);
        Assert.Equal(1045, paragraph.Runs.Single(run => run.Text == "Polish").LanguageId);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Default_And_Run_Language() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        WordParagraph polish = paragraph.AddText("Polish");
        polish._run!.RunProperties ??= new RunProperties();
        polish._run.RunProperties.Languages = new Languages { Val = "pl-PL" };
        paragraph.AddText("Default");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        Assert.Equal(1033, rtfDocument.Settings.DefaultLanguageId);
        RtfRun rtfRun = Assert.Single(Assert.Single(rtfDocument.Paragraphs).Runs, run => run.Text == "Polish");
        Assert.Equal(1045, rtfRun.LanguageId);
        Assert.Contains(@"\deflang1033", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\lang1045 Polish", rtf, StringComparison.Ordinal);

        WordParagraph roundTripRun = Assert.Single(roundTrip.Paragraphs, run => run.Text == "Polish");
        Assert.Equal("pl-PL", roundTripRun._run?.RunProperties?.Languages?.Val?.Value);
        Assert.Equal("en-US", GetDefaultWordLanguage(roundTrip));
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Default_And_Run_Language() {
        RtfDocument rtfDocument = RtfDocument.Create();
        rtfDocument.Settings.SetDefaultLanguage(1045);
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        paragraph.AddText("Default");
        paragraph.AddText(" English").SetLanguage(1033);

        using WordDocument word = rtfDocument.ToWordDocument();

        Assert.Equal("pl-PL", GetDefaultWordLanguage(word));
        WordParagraph english = Assert.Single(word.Paragraphs, run => run.Text == " English");
        Assert.Equal("en-US", english._run?.RunProperties?.Languages?.Val?.Value);
    }

    [Fact]
    public void Read_Binds_Run_Direction_Controls_Without_Leaking() {
        const string rtf = @"{\rtf1\ansi\pard \rtlch RTL\ltrch LTR\plain Plain\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal(RtfTextDirection.RightToLeft, paragraph.Runs.Single(run => run.Text == "RTL").Direction);
        Assert.Equal(RtfTextDirection.LeftToRight, paragraph.Runs.Single(run => run.Text == "LTR").Direction);
        Assert.Null(paragraph.Runs.Single(run => run.Text == "Plain").Direction);
    }

    [Fact]
    public void Write_And_Read_Run_Direction_Controls_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("RTL").SetDirection(RtfTextDirection.RightToLeft);
        paragraph.AddText("LTR").SetDirection(RtfTextDirection.LeftToRight);
        paragraph.AddText("Plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\rtlch RTL", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ltrch LTR", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\plain Plain", rtf, StringComparison.Ordinal);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal(RtfTextDirection.RightToLeft, readParagraph.Runs.Single(run => run.Text == "RTL").Direction);
        Assert.Equal(RtfTextDirection.LeftToRight, readParagraph.Runs.Single(run => run.Text == "LTR").Direction);
        Assert.Null(readParagraph.Runs.Single(run => run.Text == "Plain").Direction);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Run_Direction() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        WordParagraph rightToLeft = paragraph.AddText("RTL");
        rightToLeft._run!.RunProperties ??= new RunProperties();
        rightToLeft._run.RunProperties.RightToLeftText = new RightToLeftText();
        paragraph.AddText("Plain");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal(RtfTextDirection.RightToLeft, rtfParagraph.Runs.Single(run => run.Text == "RTL").Direction);
        Assert.Null(rtfParagraph.Runs.Single(run => run.Text == "Plain").Direction);
        Assert.Contains(@"\rtlch RTL", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\plain Plain", rtf, StringComparison.Ordinal);

        WordParagraph roundTripRtl = Assert.Single(roundTrip.Paragraphs, run => run.Text == "RTL");
        Assert.NotNull(roundTripRtl._run?.RunProperties?.RightToLeftText);
        WordParagraph roundTripPlain = Assert.Single(roundTrip.Paragraphs, run => run.Text == "Plain");
        Assert.Null(roundTripPlain._run?.RunProperties?.RightToLeftText);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Run_Direction() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        paragraph.AddText("RTL").SetDirection(RtfTextDirection.RightToLeft);
        paragraph.AddText("Plain");

        using WordDocument word = rtfDocument.ToWordDocument();

        WordParagraph rightToLeft = Assert.Single(word.Paragraphs, run => run.Text == "RTL");
        Assert.NotNull(rightToLeft._run?.RunProperties?.RightToLeftText);
        WordParagraph plain = Assert.Single(word.Paragraphs, run => run.Text == "Plain");
        Assert.Null(plain._run?.RunProperties?.RightToLeftText);
    }

    private static string? GetDefaultWordLanguage(WordDocument document) {
        return document._wordprocessingDocument.MainDocumentPart?
            .StyleDefinitionsPart?
            .Styles?
            .DocDefaults?
            .RunPropertiesDefault?
            .RunPropertiesBaseStyle?
            .Languages?
            .Val?
            .Value;
    }
}
