using OfficeIMO.Rtf;
using OfficeIMO.Html;
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

        string html = document.ToHtml(new RtfToHtmlOptions { FragmentOnly = false, NewLine = "\n" });

        Assert.Contains("<meta name=\"officeimo-rtf-styles\" content=\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-style-id=\"7\" data-officeimo-rtf-style-kind=\"paragraph\"", html, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rtf-style-id=\"8\" data-officeimo-rtf-style-kind=\"character\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
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

    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Table_Stylesheet_Row_And_Cell_Formatting() {
        RtfDocument document = RtfDocument.Create();
        int red = document.AddColor(255, 0, 0);
        int blue = document.AddColor(0, 0, 255);

        RtfStyle style = document.AddStyle(9, "Clinical Table Grid", RtfStyleKind.Table);
        RtfTableRow row = style.TableRowFormat;
        row.RepeatHeader = true;
        row.KeepTogether = true;
        row.KeepWithNext = true;
        row.SetAutoFit(false)
            .SetDirection(RtfTableRowDirection.RightToLeft)
            .SetCellGap(120)
            .SetLeftIndent(720)
            .SetAlignment(RtfTableAlignment.Center)
            .SetShading(red, blue, patternValue: 5, patternPercent: 6250, RtfShadingPattern.DarkHorizontal)
            .SetPadding(topTwips: 120, leftTwips: 180, bottomTwips: 60, rightTwips: 90)
            .SetSpacing(topTwips: 20, leftTwips: 30, bottomTwips: 40, rightTwips: 50)
            .SetPositionAnchors(RtfTableHorizontalAnchor.Page, RtfTableVerticalAnchor.Paragraph)
            .SetPosition(RtfTableHorizontalPosition.Absolute, horizontalTwips: 1440, RtfTableVerticalPosition.Bottom)
            .SetTextWrapDistances(leftTwips: 80, rightTwips: 90, topTwips: 100, bottomTwips: 110);
        row.HeightTwips = 360;
        row.PreferredWidthUnit = RtfTableWidthUnit.Twips;
        row.PreferredWidth = 5000;
        row.NoOverlap = true;
        row.TopBorder.Style = RtfTableCellBorderStyle.Single;
        row.TopBorder.Width = 12;
        row.TopBorder.ColorIndex = red;
        row.VerticalBorder.Style = RtfTableCellBorderStyle.Dotted;
        row.VerticalBorder.Width = 4;
        row.VerticalBorder.ColorIndex = blue;

        RtfTableCell cell = row.AddCell(2400);
        cell.SetPreferredWidth(2400, RtfTableWidthUnit.Twips)
            .SetNoWrap()
            .SetFitText()
            .SetHideCellMark()
            .SetShading(red, blue, patternPercent: 3750, RtfShadingPattern.DarkHorizontal)
            .SetTextFlow(RtfTableCellTextFlow.LeftToRightTopToBottomVertical)
            .SetPadding(topTwips: 60, leftTwips: 70, bottomTwips: 80, rightTwips: 90);
        cell.HorizontalMerge = RtfTableCellMerge.First;
        cell.VerticalMerge = RtfTableCellMerge.Continue;
        cell.VerticalAlignment = RtfTableCellVerticalAlignment.Center;
        cell.TopBorder.Style = RtfTableCellBorderStyle.Double;
        cell.TopBorder.Width = 8;
        cell.TopBorder.ColorIndex = blue;
        cell.TopLeftToBottomRightBorder.Style = RtfTableCellBorderStyle.Dashed;
        cell.TopLeftToBottomRightBorder.Width = 6;
        cell.TopLeftToBottomRightBorder.ColorIndex = red;

        string html = document.ToHtml(new RtfToHtmlOptions { FragmentOnly = false, NewLine = "\n" });

        Assert.Contains("<meta name=\"officeimo-rtf-styles\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.LoadFromHtml();
        RtfStyle roundTripStyle = Assert.Single(roundTrip.Styles);
        Assert.Equal(style.Id, roundTripStyle.Id);
        Assert.Equal(RtfStyleKind.Table, roundTripStyle.Kind);
        Assert.Equal("Clinical Table Grid", roundTripStyle.Name);

        RtfTableRow roundTripRow = roundTripStyle.TableRowFormat;
        Assert.True(roundTripRow.RepeatHeader);
        Assert.True(roundTripRow.KeepTogether);
        Assert.True(roundTripRow.KeepWithNext);
        Assert.Equal(false, roundTripRow.AutoFit);
        Assert.Equal(RtfTableRowDirection.RightToLeft, roundTripRow.Direction);
        Assert.Equal(360, roundTripRow.HeightTwips);
        Assert.Equal(120, roundTripRow.CellGapTwips);
        Assert.Equal(720, roundTripRow.LeftIndentTwips);
        Assert.Equal(RtfTableAlignment.Center, roundTripRow.Alignment);
        Assert.Equal(RtfTableWidthUnit.Twips, roundTripRow.PreferredWidthUnit);
        Assert.Equal(5000, roundTripRow.PreferredWidth);
        Assert.Equal(red, roundTripRow.BackgroundColorIndex);
        Assert.Equal(blue, roundTripRow.ShadingForegroundColorIndex);
        Assert.Equal(5, roundTripRow.ShadingPatternValue);
        Assert.Equal(6250, roundTripRow.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, roundTripRow.ShadingPattern);
        Assert.Equal(120, roundTripRow.PaddingTopTwips);
        Assert.Equal(180, roundTripRow.PaddingLeftTwips);
        Assert.Equal(60, roundTripRow.PaddingBottomTwips);
        Assert.Equal(90, roundTripRow.PaddingRightTwips);
        Assert.Equal(20, roundTripRow.SpacingTopTwips);
        Assert.Equal(30, roundTripRow.SpacingLeftTwips);
        Assert.Equal(40, roundTripRow.SpacingBottomTwips);
        Assert.Equal(50, roundTripRow.SpacingRightTwips);
        Assert.True(roundTripRow.NoOverlap);
        Assert.Equal(RtfTableHorizontalAnchor.Page, roundTripRow.HorizontalAnchor);
        Assert.Equal(RtfTableVerticalAnchor.Paragraph, roundTripRow.VerticalAnchor);
        Assert.Equal(RtfTableHorizontalPosition.Absolute, roundTripRow.HorizontalPosition);
        Assert.Equal(1440, roundTripRow.HorizontalPositionTwips);
        Assert.Equal(RtfTableVerticalPosition.Bottom, roundTripRow.VerticalPosition);
        Assert.Equal(80, roundTripRow.TextWrapLeftTwips);
        Assert.Equal(90, roundTripRow.TextWrapRightTwips);
        Assert.Equal(100, roundTripRow.TextWrapTopTwips);
        Assert.Equal(110, roundTripRow.TextWrapBottomTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, roundTripRow.TopBorder.Style);
        Assert.Equal(12, roundTripRow.TopBorder.Width);
        Assert.Equal(red, roundTripRow.TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, roundTripRow.VerticalBorder.Style);
        Assert.Equal(4, roundTripRow.VerticalBorder.Width);
        Assert.Equal(blue, roundTripRow.VerticalBorder.ColorIndex);

        RtfTableCell roundTripCell = Assert.Single(roundTripRow.Cells);
        Assert.Equal(2400, roundTripCell.RightBoundaryTwips);
        Assert.Equal(RtfTableCellMerge.First, roundTripCell.HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, roundTripCell.VerticalMerge);
        Assert.Equal(RtfTableWidthUnit.Twips, roundTripCell.PreferredWidthUnit);
        Assert.Equal(2400, roundTripCell.PreferredWidth);
        Assert.True(roundTripCell.HideCellMark);
        Assert.True(roundTripCell.NoWrap);
        Assert.True(roundTripCell.FitText);
        Assert.Equal(red, roundTripCell.BackgroundColorIndex);
        Assert.Equal(blue, roundTripCell.ShadingForegroundColorIndex);
        Assert.Equal(3750, roundTripCell.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, roundTripCell.ShadingPattern);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, roundTripCell.VerticalAlignment);
        Assert.Equal(RtfTableCellTextFlow.LeftToRightTopToBottomVertical, roundTripCell.TextFlow);
        Assert.Equal(60, roundTripCell.PaddingTopTwips);
        Assert.Equal(70, roundTripCell.PaddingLeftTwips);
        Assert.Equal(80, roundTripCell.PaddingBottomTwips);
        Assert.Equal(90, roundTripCell.PaddingRightTwips);
        Assert.Equal(RtfTableCellBorderStyle.Double, roundTripCell.TopBorder.Style);
        Assert.Equal(8, roundTripCell.TopBorder.Width);
        Assert.Equal(blue, roundTripCell.TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, roundTripCell.TopLeftToBottomRightBorder.Style);
        Assert.Equal(6, roundTripCell.TopLeftToBottomRightBorder.Width);
        Assert.Equal(red, roundTripCell.TopLeftToBottomRightBorder.ColorIndex);

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\*\ts9\tsrowd", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trhdr", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trkeepfollow", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbgdkhor", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clNoWrap\clFitText", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cldglu", rtf, StringComparison.Ordinal);
    }
}
