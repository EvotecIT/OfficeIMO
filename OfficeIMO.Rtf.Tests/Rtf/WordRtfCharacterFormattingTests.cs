using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class WordRtfCharacterFormattingTests {
    [Fact]
    public void Word_Rtf_Bridge_RoundTrips_Superscript_And_Subscript_Through_Core_Model() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("2");
        paragraph.AddText("nd").SetSuperScript();
        paragraph.AddText(" H");
        paragraph.AddText("2").SetSubScript();
        paragraph.AddText("O");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Equal("2nd H2O", rtfParagraph.ToPlainText());
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "nd" && run.VerticalPosition == RtfVerticalPosition.Superscript);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "2" && run.VerticalPosition == RtfVerticalPosition.Subscript);
        Assert.Contains(@"\super nd\nosupersub", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\sub 2\nosupersub", rtf, StringComparison.Ordinal);

        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "nd" && run.VerticalTextAlignment == VerticalPositionValues.Superscript);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "2" && run.VerticalTextAlignment == VerticalPositionValues.Subscript);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Hidden_Text() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("Visible ");
        WordParagraph hidden = paragraph.AddText("Hidden");
        hidden._run!.RunProperties ??= new RunProperties();
        hidden._run.RunProperties.Vanish = new Vanish();
        paragraph.AddText(" shown");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Hidden" && run.Hidden);
        Assert.Contains(@"\v Hidden\v0", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Hidden" && run._run?.RunProperties?.Vanish != null);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Double_Strike_And_Caps_Effects() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("Double").SetDoubleStrike();
        paragraph.AddText("Caps").SetCapsStyle(CapsStyle.Caps);
        paragraph.AddText("Small").SetSmallCaps();
        paragraph.AddText("Plain");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Double" && run.DoubleStrike);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Caps" && run.CapsStyle == RtfCapsStyle.Caps);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Small" && run.CapsStyle == RtfCapsStyle.SmallCaps);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Plain" && !run.DoubleStrike && run.CapsStyle == RtfCapsStyle.None);
        Assert.Contains(@"\striked Double\striked0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\caps Caps\caps0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\scaps Small\scaps0", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Double" && run.DoubleStrike);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Caps" && run.CapsStyle == CapsStyle.Caps);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Small" && run.CapsStyle == CapsStyle.SmallCaps);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Plain" && !run.DoubleStrike && run.CapsStyle == CapsStyle.None);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Double_Strike_And_Caps_Effects() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        paragraph.AddText("Double").SetDoubleStrike();
        paragraph.AddText("Caps").SetCapsStyle(RtfCapsStyle.Caps);
        paragraph.AddText("Small").SetCapsStyle(RtfCapsStyle.SmallCaps);
        paragraph.AddText("Plain");

        using WordDocument word = rtfDocument.ToWordDocument();

        Assert.Contains(word.Paragraphs, run => run.Text == "Double" && run.DoubleStrike);
        Assert.Contains(word.Paragraphs, run => run.Text == "Caps" && run.CapsStyle == CapsStyle.Caps);
        Assert.Contains(word.Paragraphs, run => run.Text == "Small" && run.CapsStyle == CapsStyle.SmallCaps);
        Assert.Contains(word.Paragraphs, run => run.Text == "Plain" && !run.DoubleStrike && run.CapsStyle == CapsStyle.None);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Outline_Shadow_Emboss_And_Imprint_Effects() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("Outline").SetOutline();
        paragraph.AddText("Shadow").SetShadow();
        paragraph.AddText("Emboss").SetEmboss();
        WordParagraph imprint = paragraph.AddText("Imprint");
        imprint._run!.RunProperties ??= new RunProperties();
        imprint._run.RunProperties.Imprint = new Imprint();
        paragraph.AddText("Plain");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Outline" && run.Outline);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Shadow" && run.Shadow);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Emboss" && run.Emboss);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Imprint" && run.Imprint);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Plain" && !run.Outline && !run.Shadow && !run.Emboss && !run.Imprint);
        Assert.Contains(@"\outl Outline\outl0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\shad Shadow\shad0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\embo Emboss\embo0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\impr Imprint\impr0", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Outline" && run.Outline);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Shadow" && run.Shadow);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Emboss" && run.Emboss);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Imprint" && run._run?.RunProperties?.Imprint != null);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Plain" && !run.Outline && !run.Shadow && !run.Emboss && run._run?.RunProperties?.Imprint == null);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Outline_Shadow_Emboss_And_Imprint_Effects() {
        RtfDocument rtfDocument = RtfDocument.Create();
        RtfParagraph paragraph = rtfDocument.AddParagraph();
        paragraph.AddText("Outline").SetOutline();
        paragraph.AddText("Shadow").SetShadow();
        paragraph.AddText("Emboss").SetEmboss();
        paragraph.AddText("Imprint").SetImprint();
        paragraph.AddText("Plain");

        using WordDocument word = rtfDocument.ToWordDocument();

        Assert.Contains(word.Paragraphs, run => run.Text == "Outline" && run.Outline);
        Assert.Contains(word.Paragraphs, run => run.Text == "Shadow" && run.Shadow);
        Assert.Contains(word.Paragraphs, run => run.Text == "Emboss" && run.Emboss);
        Assert.Contains(word.Paragraphs, run => run.Text == "Imprint" && run._run?.RunProperties?.Imprint != null);
        Assert.Contains(word.Paragraphs, run => run.Text == "Plain" && !run.Outline && !run.Shadow && !run.Emboss && run._run?.RunProperties?.Imprint == null);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Run_Highlight() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("Normal ");
        paragraph.AddText("Marked").SetHighlight(HighlightColorValues.Yellow);
        paragraph.AddText(" done");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Contains(rtfParagraph.Runs, run => run.Text == "Marked" && run.HighlightColorIndex == 1);
        Assert.Contains(@"{\colortbl;\red255\green255\blue0;}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\highlight1 Marked", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Marked" && run.Highlight == HighlightColorValues.Yellow);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Run_Highlight() {
        RtfDocument rtfDocument = RtfDocument.Create();
        int yellow = rtfDocument.AddColor(255, 255, 0);
        RtfParagraph paragraph = rtfDocument.AddParagraph("Normal ");
        paragraph.AddText("Marked").SetHighlightColor(yellow);
        paragraph.AddText(" done");

        using WordDocument word = rtfDocument.ToWordDocument();

        Assert.Contains(word.Paragraphs, run => run.Text == "Marked" && run.Highlight == HighlightColorValues.Yellow);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Run_Font_And_Color() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("Normal ");
        paragraph.AddText("Styled").SetFontFamily("Consolas").SetColorHex("4472C4");
        paragraph.AddText(" done");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = RtfDocument.Read(rtf).Document.ToWordDocument();

        Assert.Contains(rtfDocument.Fonts, font => font.Id == 1 && font.Name == "Consolas");
        Assert.Contains(rtfDocument.Colors, color => color.Red == 0x44 && color.Green == 0x72 && color.Blue == 0xC4);
        RtfRun styledRun = Assert.Single(Assert.Single(rtfDocument.Paragraphs).Runs, run => run.Text == "Styled");
        Assert.Equal(1, styledRun.FontId);
        Assert.Equal(1, styledRun.ForegroundColorIndex);
        Assert.Contains(@"{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\colortbl;\red68\green114\blue196;}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\f1 \cf1 Styled", rtf, StringComparison.Ordinal);
        Assert.Contains(roundTrip.Paragraphs, run => run.Text == "Styled" && run.FontFamily == "Consolas" && run.ColorHex == "4472C4");
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Run_Font_And_Color() {
        RtfDocument rtfDocument = RtfDocument.Create();
        int fontId = rtfDocument.AddFont("Consolas");
        int colorId = rtfDocument.AddColor(0x44, 0x72, 0xC4);
        RtfParagraph paragraph = rtfDocument.AddParagraph("Normal ");
        RtfRun styled = paragraph.AddText("Styled");
        styled.FontId = fontId;
        styled.ForegroundColorIndex = colorId;
        paragraph.AddText(" done");

        using WordDocument word = rtfDocument.ToWordDocument();

        Assert.Contains(word.Paragraphs, run => run.Text == "Styled" && run.FontFamily == "Consolas" && run.ColorHex == "4472C4");
    }
}
