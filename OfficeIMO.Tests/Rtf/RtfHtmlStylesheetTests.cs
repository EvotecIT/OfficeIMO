using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlStylesheetTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Stylesheet_And_Style_References() {
        RtfDocument document = RtfDocument.Create();
        int fontId = document.AddFont("Consolas");
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);
        int yellow = document.AddColor(255, 255, 0);

        RtfStyle heading = document.AddStyle(7, "Clinical Heading");
        heading.KeyCode = new RtfStyleKeyCode {
            Shift = true,
            Control = true,
            FunctionKey = 3,
            Key = "h"
        };
        heading.BasedOnStyleId = 0;
        heading.NextStyleId = 7;
        heading.LinkedStyleId = 8;
        heading.AutoUpdate = true;
        heading.Hidden = true;
        heading.Locked = true;
        heading.Personal = true;
        heading.Compose = true;
        heading.Reply = true;
        heading.SemiHidden = true;
        heading.UnhideWhenUsed = true;
        heading.QuickFormat = true;
        heading.Priority = 9;
        heading.RevisionSaveId = 123;
        heading.Bold = true;
        heading.Italic = false;
        heading.UnderlineStyle = RtfUnderlineStyle.Double;
        heading.FontSize = 15.5;
        heading.FontId = fontId;
        heading.ForegroundColorIndex = red;
        heading.HighlightColorIndex = yellow;
        heading.ParagraphAlignment = RtfTextAlignment.Justify;
        heading.ParagraphDirection = RtfTextDirection.RightToLeft;
        heading.LeftIndentTwips = 720;
        heading.RightIndentTwips = 360;
        heading.FirstLineIndentTwips = -180;
        heading.SpaceBeforeTwips = 120;
        heading.SpaceAfterTwips = 240;
        heading.SpaceBeforeAuto = false;
        heading.SpaceAfterAuto = true;
        heading.LineSpacingTwips = 360;
        heading.LineSpacingMultiple = true;
        heading.BackgroundColorIndex = red;
        heading.ShadingForegroundColorIndex = blue;
        heading.ShadingPatternPercent = 5000;
        heading.ShadingPattern = RtfShadingPattern.DarkHorizontal;
        heading.PageBreakBefore = true;
        heading.KeepWithNext = true;
        heading.KeepLinesTogether = true;
        heading.SuppressLineNumbers = true;
        heading.AutoHyphenation = false;
        heading.ContextualSpacing = false;
        heading.AdjustRightIndent = true;
        heading.SnapToLineGrid = false;
        heading.WidowControl = false;
        heading.OutlineLevel = 2;
        heading.SetBorder(RtfParagraphBorderSide.Top, RtfParagraphBorderStyle.Single, width: 12, colorIndex: red)
            .SetBorder(RtfParagraphBorderSide.Left, RtfParagraphBorderStyle.Double, width: 8, colorIndex: blue);
        heading.AddTabStop(2880, RtfTabAlignment.Right, RtfTabLeader.Dots);

        RtfStyle character = document.AddStyle(8, "Clinical Link", RtfStyleKind.Character);
        character.Additive = true;
        character.LinkedStyleId = 7;
        character.Italic = true;

        RtfParagraph paragraph = document.AddParagraph();
        paragraph.SetStyle(heading.Id);
        paragraph.AddText("Clinical heading").SetStyle(character.Id);

        string html = document.ToHtml(new RtfHtmlSaveOptions { FragmentOnly = false, NewLine = "\n" });

        Assert.Contains("<meta name=\"officeimo-rtf-styles\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-style-id=\"7\" data-officeimo-rtf-style-kind=\"paragraph\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-style-id=\"8\" data-officeimo-rtf-style-kind=\"character\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocumentFromHtml();
        RtfStyle roundTripHeading = roundTrip.Styles.Single(style => style.Id == 7 && style.Kind == RtfStyleKind.Paragraph);
        Assert.Equal("Clinical Heading", roundTripHeading.Name);
        Assert.Equal(0, roundTripHeading.BasedOnStyleId);
        Assert.Equal(7, roundTripHeading.NextStyleId);
        Assert.Equal(8, roundTripHeading.LinkedStyleId);
        Assert.NotNull(roundTripHeading.KeyCode);
        Assert.True(roundTripHeading.KeyCode.Shift);
        Assert.True(roundTripHeading.KeyCode.Control);
        Assert.Equal(3, roundTripHeading.KeyCode.FunctionKey);
        Assert.Equal("h", roundTripHeading.KeyCode.Key);
        Assert.True(roundTripHeading.AutoUpdate);
        Assert.True(roundTripHeading.Hidden);
        Assert.True(roundTripHeading.Locked);
        Assert.True(roundTripHeading.Personal);
        Assert.True(roundTripHeading.Compose);
        Assert.True(roundTripHeading.Reply);
        Assert.True(roundTripHeading.SemiHidden);
        Assert.True(roundTripHeading.UnhideWhenUsed);
        Assert.True(roundTripHeading.QuickFormat);
        Assert.Equal(9, roundTripHeading.Priority);
        Assert.Equal(123, roundTripHeading.RevisionSaveId);
        Assert.Equal(true, roundTripHeading.Bold);
        Assert.Equal(false, roundTripHeading.Italic);
        Assert.Equal(RtfUnderlineStyle.Double, roundTripHeading.UnderlineStyle);
        Assert.Equal(15.5, roundTripHeading.FontSize);
        Assert.Equal(fontId, roundTripHeading.FontId);
        Assert.Equal(red, roundTripHeading.ForegroundColorIndex);
        Assert.Equal(yellow, roundTripHeading.HighlightColorIndex);
        Assert.Equal(RtfTextAlignment.Justify, roundTripHeading.ParagraphAlignment);
        Assert.Equal(RtfTextDirection.RightToLeft, roundTripHeading.ParagraphDirection);
        Assert.Equal(720, roundTripHeading.LeftIndentTwips);
        Assert.Equal(360, roundTripHeading.RightIndentTwips);
        Assert.Equal(-180, roundTripHeading.FirstLineIndentTwips);
        Assert.Equal(120, roundTripHeading.SpaceBeforeTwips);
        Assert.Equal(240, roundTripHeading.SpaceAfterTwips);
        Assert.Equal(false, roundTripHeading.SpaceBeforeAuto);
        Assert.Equal(true, roundTripHeading.SpaceAfterAuto);
        Assert.Equal(360, roundTripHeading.LineSpacingTwips);
        Assert.Equal(true, roundTripHeading.LineSpacingMultiple);
        Assert.Equal(red, roundTripHeading.BackgroundColorIndex);
        Assert.Equal(blue, roundTripHeading.ShadingForegroundColorIndex);
        Assert.Equal(5000, roundTripHeading.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, roundTripHeading.ShadingPattern);
        Assert.Equal(true, roundTripHeading.PageBreakBefore);
        Assert.Equal(true, roundTripHeading.KeepWithNext);
        Assert.Equal(true, roundTripHeading.KeepLinesTogether);
        Assert.Equal(true, roundTripHeading.SuppressLineNumbers);
        Assert.Equal(false, roundTripHeading.AutoHyphenation);
        Assert.Equal(false, roundTripHeading.ContextualSpacing);
        Assert.Equal(true, roundTripHeading.AdjustRightIndent);
        Assert.Equal(false, roundTripHeading.SnapToLineGrid);
        Assert.Equal(false, roundTripHeading.WidowControl);
        Assert.Equal(2, roundTripHeading.OutlineLevel);
        Assert.Equal(RtfParagraphBorderStyle.Single, roundTripHeading.TopBorder.Style);
        Assert.Equal(12, roundTripHeading.TopBorder.Width);
        Assert.Equal(red, roundTripHeading.TopBorder.ColorIndex);
        Assert.Equal(RtfParagraphBorderStyle.Double, roundTripHeading.LeftBorder.Style);
        Assert.Equal(8, roundTripHeading.LeftBorder.Width);
        Assert.Equal(blue, roundTripHeading.LeftBorder.ColorIndex);
        RtfTabStop tabStop = Assert.Single(roundTripHeading.TabStops);
        Assert.Equal(2880, tabStop.PositionTwips);
        Assert.Equal(RtfTabAlignment.Right, tabStop.Alignment);
        Assert.Equal(RtfTabLeader.Dots, tabStop.Leader);

        RtfStyle roundTripCharacter = roundTrip.Styles.Single(style => style.Id == 8 && style.Kind == RtfStyleKind.Character);
        Assert.Equal("Clinical Link", roundTripCharacter.Name);
        Assert.True(roundTripCharacter.Additive);
        Assert.Equal(7, roundTripCharacter.LinkedStyleId);
        Assert.Equal(true, roundTripCharacter.Italic);
        Assert.Equal(7, Assert.Single(roundTrip.Paragraphs).StyleId);
        Assert.Equal(8, Assert.Single(roundTrip.Paragraphs[0].Runs).StyleId);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\s7", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\*\cs8", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\slink8", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tldot\tqr\tx2880", rtf, StringComparison.Ordinal);
    }
}
