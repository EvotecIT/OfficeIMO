using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Word_Rtf_Bridge_Carries_Tables_As_Table_Blocks_In_Document_Order() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Before");
        WordTable table = word.AddTable(2, 2);
        table.CheckTableProperties();
        table._tableProperties!.TableCellSpacing = new TableCellSpacing { Width = "300", Type = TableWidthUnitValues.Dxa };
        table._tableProperties.TableIndentation = new TableIndentation { Width = 720, Type = TableWidthUnitValues.Dxa };
        table.Rows[0].Cells[0].AddParagraph("A1", removeExistingParagraphs: true);
        table.Rows[0].Cells[1].AddParagraph("B1", removeExistingParagraphs: true);
        table.Rows[1].Cells[0].AddParagraph("A2", removeExistingParagraphs: true);
        table.Rows[1].Cells[1].AddParagraph("B2", removeExistingParagraphs: true);
        table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
        table.Rows[0].Cells[0].Width = 1800;
        table.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;
        table.Rows[0].Cells[1].Width = 3000;
        word.AddParagraph("After");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        Assert.Collection(rtfDocument.Blocks,
            block => Assert.Equal("Before", Assert.IsType<RtfParagraph>(block).ToPlainText()),
            block => {
                RtfTable rtfTable = Assert.IsType<RtfTable>(block);
                Assert.Equal(2, rtfTable.Rows.Count);
                Assert.Equal("A1", rtfTable.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
                Assert.Equal("B2", rtfTable.Rows[1].Cells[1].Paragraphs[0].ToPlainText());
                Assert.Equal(300, rtfTable.Rows[0].CellGapTwips);
                Assert.Equal(300, rtfTable.Rows[1].CellGapTwips);
                Assert.Equal(720, rtfTable.Rows[0].LeftIndentTwips);
                Assert.Equal(720, rtfTable.Rows[1].LeftIndentTwips);
                Assert.Equal(1800, rtfTable.Rows[0].Cells[0].RightBoundaryTwips);
                Assert.Equal(4800, rtfTable.Rows[0].Cells[1].RightBoundaryTwips);
                Assert.Equal(RtfTableWidthUnit.Twips, rtfTable.Rows[0].Cells[0].PreferredWidthUnit);
                Assert.Equal(1800, rtfTable.Rows[0].Cells[0].PreferredWidth);
                Assert.Equal(RtfTableWidthUnit.Twips, rtfTable.Rows[0].Cells[1].PreferredWidthUnit);
                Assert.Equal(3000, rtfTable.Rows[0].Cells[1].PreferredWidth);
            },
            block => Assert.Equal("After", Assert.IsType<RtfParagraph>(block).ToPlainText()));
        Assert.Contains(@"\trowd", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trgaph300", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trleft720", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth3\clwWidth1800\cellx1800", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth3\clwWidth3000\cellx4800", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cellx1800", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cellx4800", rtf, StringComparison.Ordinal);

        WordTable roundTripTable = Assert.Single(roundTrip.Tables);
        Assert.Equal("300", roundTripTable._tableProperties?.TableCellSpacing?.Width?.Value);
        Assert.Equal(720, roundTripTable._tableProperties?.TableIndentation?.Width?.Value);
        Assert.Equal(TableWidthUnitValues.Dxa, roundTripTable.Rows[0].Cells[0].WidthType);
        Assert.Equal(1800, roundTripTable.Rows[0].Cells[0].Width);
        Assert.Equal(TableWidthUnitValues.Dxa, roundTripTable.Rows[0].Cells[1].WidthType);
        Assert.Equal(3000, roundTripTable.Rows[0].Cells[1].Width);
        Assert.Equal("A1", GetCellText(roundTripTable, 0, 0));
        Assert.Equal("B1", GetCellText(roundTripTable, 0, 1));
        Assert.Equal("A2", GetCellText(roundTripTable, 1, 0));
        Assert.Equal("B2", GetCellText(roundTripTable, 1, 1));
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Table_Blocks_As_Word_Tables() {
        RtfDocument rtfDocument = RtfDocument.Create();
        rtfDocument.AddParagraph("Before");
        RtfTable table = rtfDocument.AddTable(0, 1);
        RtfTableRow row = table.AddRow();
        row.SetCellGap(300);
        row.SetLeftIndent(720);
        RtfTableCell firstCell = row.AddCell(2400);
        firstCell.SetPreferredWidth(1800, RtfTableWidthUnit.Twips);
        firstCell.SetNoWrap();
        firstCell.SetFitText();
        RtfParagraph firstParagraph = firstCell.AddParagraph("Left ");
        firstParagraph.AddText("bold").SetBold();
        RtfTableCell secondCell = row.AddCell(4800);
        secondCell.SetPreferredWidth(3000, RtfTableWidthUnit.Twips);
        secondCell.AddParagraph("Right");
        rtfDocument.AddParagraph("After");

        using WordDocument word = rtfDocument.ToWordDocument();
        RtfDocument roundTrip = word.ToRtfDocument();

        WordTable wordTable = Assert.Single(word.Tables);
        Assert.Equal("300", wordTable._tableProperties?.TableCellSpacing?.Width?.Value);
        Assert.Equal(720, wordTable._tableProperties?.TableIndentation?.Width?.Value);
        Assert.Equal(TableWidthUnitValues.Dxa, wordTable.Rows[0].Cells[0].WidthType);
        Assert.Equal(1800, wordTable.Rows[0].Cells[0].Width);
        Assert.False(wordTable.Rows[0].Cells[0].WrapText);
        Assert.True(wordTable.Rows[0].Cells[0].FitText);
        Assert.Equal(TableWidthUnitValues.Dxa, wordTable.Rows[0].Cells[1].WidthType);
        Assert.Equal(3000, wordTable.Rows[0].Cells[1].Width);
        Assert.Equal("Left bold", GetCellText(wordTable, 0, 0));
        Assert.Equal("Right", GetCellText(wordTable, 0, 1));
        Assert.Contains(wordTable.Rows[0].Cells[0].Paragraphs, paragraph => paragraph.Text == "bold" && paragraph.Bold);
        Assert.Collection(roundTrip.Blocks,
            block => Assert.Equal("Before", Assert.IsType<RtfParagraph>(block).ToPlainText()),
            block => {
                RtfTable roundTripTable = Assert.IsType<RtfTable>(block);
                Assert.Equal(300, roundTripTable.Rows[0].CellGapTwips);
                Assert.Equal(720, roundTripTable.Rows[0].LeftIndentTwips);
                Assert.Equal(RtfTableWidthUnit.Twips, roundTripTable.Rows[0].Cells[0].PreferredWidthUnit);
                Assert.Equal(1800, roundTripTable.Rows[0].Cells[0].PreferredWidth);
                Assert.True(roundTripTable.Rows[0].Cells[0].NoWrap);
                Assert.True(roundTripTable.Rows[0].Cells[0].FitText);
                Assert.Equal(RtfTableWidthUnit.Twips, roundTripTable.Rows[0].Cells[1].PreferredWidthUnit);
                Assert.Equal(3000, roundTripTable.Rows[0].Cells[1].PreferredWidth);
                Assert.Equal("Left bold", roundTripTable.Rows[0].Cells[0].Paragraphs[0].ToPlainText());
            },
            block => Assert.Equal("After", Assert.IsType<RtfParagraph>(block).ToPlainText()));
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Embedded_Picture_Blocks() {
        byte[] png = CreateOnePixelPng();
        using WordDocument word = WordDocument.Create();
        using (var stream = new MemoryStream(png)) {
            word.AddParagraph().AddImage(stream, "pixel.png", 32, 16, WrapTextImage.InLineWithText, "Pixel image");
        }

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();
        RtfDocument roundTripRtf = roundTrip.ToRtfDocument();

        RtfImage image = Assert.IsType<RtfImage>(Assert.Single(rtfDocument.Blocks));
        Assert.Equal(RtfImageFormat.Png, image.Format);
        Assert.Equal(png, image.Data);
        Assert.Equal(32, image.SourceWidth);
        Assert.Equal(16, image.SourceHeight);
        Assert.Equal(480, image.DesiredWidthTwips);
        Assert.Equal(240, image.DesiredHeightTwips);
        Assert.Equal("Pixel image", image.Description);
        Assert.Contains(@"{\pict\pngblip\picw32\pich16\picwgoal480\pichgoal240", rtf, StringComparison.Ordinal);
        Assert.Contains("89504e470d0a1a0a", rtf, StringComparison.Ordinal);

        using WordDocument semanticBridge = rtfDocument.ToWordDocument();
        Assert.Equal("Pixel image", Assert.Single(semanticBridge.Images).Description);
        WordImage roundTripImage = Assert.Single(roundTrip.Images);
        Assert.Equal(png, roundTripImage.ToBytes());
        RtfImage semanticRoundTripImage = Assert.IsType<RtfImage>(Assert.Single(roundTripRtf.Blocks));
        Assert.Equal(png, semanticRoundTripImage.Data);
        Assert.Equal(480, semanticRoundTripImage.DesiredWidthTwips);
        Assert.Equal(240, semanticRoundTripImage.DesiredHeightTwips);
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Embedded_Pictures_In_Paragraph_Order() {
        byte[] png = CreateOnePixelPng();
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph.AddText("Before ");
        using (var stream = new MemoryStream(png)) {
            paragraph.AddImage(stream, "pixel.png", 24, 12, WrapTextImage.InLineWithText, "Inline pixel");
        }

        paragraph.AddText(" after");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();
        RtfDocument semanticRoundTrip = roundTrip.ToRtfDocument();

        RtfParagraph rtfParagraph = Assert.Single(rtfDocument.Paragraphs);
        Assert.Collection(rtfParagraph.Inlines,
            inline => Assert.Equal("Before ", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfImage image = Assert.IsType<RtfImage>(inline);
                Assert.Equal(RtfImageFormat.Png, image.Format);
                Assert.Equal(png, image.Data);
                Assert.Equal(24, image.SourceWidth);
                Assert.Equal(12, image.SourceHeight);
                Assert.Equal(360, image.DesiredWidthTwips);
                Assert.Equal(180, image.DesiredHeightTwips);
                Assert.Equal("Inline pixel", image.Description);
            },
            inline => Assert.Equal(" after", Assert.IsType<RtfRun>(inline).Text));
        Assert.Contains(@"{\pict\pngblip\picw24\pich12\picwgoal360\pichgoal180", rtf, StringComparison.Ordinal);
        Assert.Equal("Before  after", string.Concat(roundTrip.Paragraphs.Select(item => item.Text)));
        Assert.Single(roundTrip.Images);

        RtfParagraph roundTripParagraph = Assert.Single(semanticRoundTrip.Paragraphs);
        Assert.Collection(roundTripParagraph.Inlines,
            inline => Assert.Equal("Before ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(png, Assert.IsType<RtfImage>(inline).Data),
            inline => Assert.Equal(" after", Assert.IsType<RtfRun>(inline).Text));
    }

    [Fact]
    public void Word_Rtf_Bridge_Carries_Table_Row_And_Cell_Formatting() {
        using WordDocument word = WordDocument.Create();
        WordTable table = word.AddTable(2, 3);
        table.Alignment = TableRowAlignmentValues.Center;
        table.WidthType = TableWidthUnitValues.Pct;
        table.Width = 4250;
        table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage = true;
        table.Rows[0].AllowRowToBreakAcrossPages = false;
        table.Rows[0].Height = 640;
        table.Rows[0].Cells[0].HorizontalMerge = MergedCellValues.Restart;
        table.Rows[0].Cells[1].HorizontalMerge = MergedCellValues.Continue;
        table.Rows[0].Cells[0].VerticalMerge = MergedCellValues.Restart;
        table.Rows[1].Cells[0].VerticalMerge = MergedCellValues.Continue;
        table.Rows[0].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Center;
        table.Rows[0].Cells[0].TextDirection = TextDirectionValues.TopToBottomRightToLeftRotated;
        table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
        table.Rows[0].Cells[0].Width = 1800;
        table.Rows[0].Cells[0].WrapText = false;
        table.Rows[0].Cells[0].FitText = true;
        table.Rows[1].Cells[2].VerticalAlignment = TableVerticalAlignmentValues.Bottom;
        table.Rows[0].Cells[0].ShadingFillColorHex = "E6F2FF";
        table.Rows[0].Cells[0].ShadingPattern = ShadingPatternValues.Percent37;
        table.Rows[0].Cells[0]._tableCellProperties!.Shading!.Color = "00AA55";
        table.Rows[0].Cells[0].MarginTopWidth = 120;
        table.Rows[0].Cells[0].MarginLeftWidth = 180;
        table.Rows[0].Cells[0].MarginBottomWidth = 240;
        table.Rows[0].Cells[0].MarginRightWidth = 300;
        table.Rows[0].Cells[0].Borders.TopStyle = BorderValues.Single;
        table.Rows[0].Cells[0].Borders.TopSize = 12U;
        table.Rows[0].Cells[0].Borders.TopColorHex = "4472C4";
        table.Rows[0].Cells[0].Borders.LeftStyle = BorderValues.Double;
        table.Rows[0].Cells[0].Borders.LeftSize = 8U;
        table.Rows[0].Cells[0].Borders.LeftColorHex = "00AA55";
        table.Rows[0].Cells[0].Borders.TopLeftToBottomRightStyle = BorderValues.Dotted;
        table.Rows[0].Cells[0].Borders.TopLeftToBottomRightSize = 6U;
        table.Rows[0].Cells[0].Borders.TopLeftToBottomRightColorHex = "4472C4";
        table.Rows[0].Cells[0].Borders.TopRightToBottomLeftStyle = BorderValues.Dashed;
        table.Rows[0].Cells[0].Borders.TopRightToBottomLeftSize = 10U;
        table.Rows[0].Cells[0].Borders.TopRightToBottomLeftColorHex = "00AA55";
        table.Rows[0].Cells[0].AddParagraph("Merged", removeExistingParagraphs: true);
        table.Rows[1].Cells[0].AddParagraph("Vertical", removeExistingParagraphs: true);

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        using WordDocument roundTrip = rtf.LoadFromRtf();

        RtfTable rtfTable = Assert.IsType<RtfTable>(Assert.Single(rtfDocument.Blocks));
        Assert.Equal(RtfTableAlignment.Center, rtfTable.Rows[0].Alignment);
        Assert.Equal(RtfTableAlignment.Center, rtfTable.Rows[1].Alignment);
        Assert.Equal(RtfTableWidthUnit.Percent, rtfTable.Rows[0].PreferredWidthUnit);
        Assert.Equal(4250, rtfTable.Rows[0].PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Percent, rtfTable.Rows[1].PreferredWidthUnit);
        Assert.Equal(4250, rtfTable.Rows[1].PreferredWidth);
        Assert.True(rtfTable.Rows[0].RepeatHeader);
        Assert.True(rtfTable.Rows[0].KeepTogether);
        Assert.Equal(640, rtfTable.Rows[0].HeightTwips);
        Assert.Equal(RtfTableCellMerge.First, rtfTable.Rows[0].Cells[0].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, rtfTable.Rows[0].Cells[1].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.First, rtfTable.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, rtfTable.Rows[1].Cells[0].VerticalMerge);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, rtfTable.Rows[0].Cells[0].VerticalAlignment);
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeftVertical, rtfTable.Rows[0].Cells[0].TextFlow);
        Assert.Equal(RtfTableWidthUnit.Twips, rtfTable.Rows[0].Cells[0].PreferredWidthUnit);
        Assert.Equal(1800, rtfTable.Rows[0].Cells[0].PreferredWidth);
        Assert.True(rtfTable.Rows[0].Cells[0].NoWrap);
        Assert.True(rtfTable.Rows[0].Cells[0].FitText);
        Assert.Equal(RtfTableCellVerticalAlignment.Bottom, rtfTable.Rows[1].Cells[2].VerticalAlignment);
        Assert.Equal(1, rtfTable.Rows[0].Cells[0].BackgroundColorIndex);
        Assert.Equal(2, rtfTable.Rows[0].Cells[0].ShadingForegroundColorIndex);
        Assert.Equal(3750, rtfTable.Rows[0].Cells[0].ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.None, rtfTable.Rows[0].Cells[0].ShadingPattern);
        Assert.Equal(120, rtfTable.Rows[0].Cells[0].PaddingTopTwips);
        Assert.Equal(180, rtfTable.Rows[0].Cells[0].PaddingLeftTwips);
        Assert.Equal(240, rtfTable.Rows[0].Cells[0].PaddingBottomTwips);
        Assert.Equal(300, rtfTable.Rows[0].Cells[0].PaddingRightTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, rtfTable.Rows[0].Cells[0].TopBorder.Style);
        Assert.Equal(12, rtfTable.Rows[0].Cells[0].TopBorder.Width);
        Assert.Equal(3, rtfTable.Rows[0].Cells[0].TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Double, rtfTable.Rows[0].Cells[0].LeftBorder.Style);
        Assert.Equal(8, rtfTable.Rows[0].Cells[0].LeftBorder.Width);
        Assert.Equal(2, rtfTable.Rows[0].Cells[0].LeftBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, rtfTable.Rows[0].Cells[0].TopLeftToBottomRightBorder.Style);
        Assert.Equal(6, rtfTable.Rows[0].Cells[0].TopLeftToBottomRightBorder.Width);
        Assert.Equal(3, rtfTable.Rows[0].Cells[0].TopLeftToBottomRightBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, rtfTable.Rows[0].Cells[0].TopRightToBottomLeftBorder.Style);
        Assert.Equal(10, rtfTable.Rows[0].Cells[0].TopRightToBottomLeftBorder.Width);
        Assert.Equal(2, rtfTable.Rows[0].Cells[0].TopRightToBottomLeftBorder.ColorIndex);
        Assert.Contains(@"\trqc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trftsWidth2\trwWidth4250", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trhdr", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trkeep", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clmgf", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clmrg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvmgf", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvmrg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cltxtbrlv", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth3\clwWidth1800", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clNoWrap\clFitText", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clcbpat1\clcfpat2\clshdng3750", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadt120\clpadft3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadl180\clpadfl3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadb240\clpadfb3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadr300\clpadfr3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrt\brdrs\brdrw12\brdrcf3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrl\brdrdb\brdrw8\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cldglu\brdrdot\brdrw6\brdrcf3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cldgll\brdrdash\brdrw10\brdrcf2", rtf, StringComparison.Ordinal);

        WordTable roundTripTable = Assert.Single(roundTrip.Tables);
        Assert.Equal(TableRowAlignmentValues.Center, roundTripTable.Alignment);
        Assert.Equal(TableWidthUnitValues.Pct, roundTripTable.WidthType);
        Assert.Equal(4250, roundTripTable.Width);
        Assert.True(roundTripTable.Rows[0].RepeatHeaderRowAtTheTopOfEachPage);
        Assert.False(roundTripTable.Rows[0].AllowRowToBreakAcrossPages);
        Assert.Equal(640, roundTripTable.Rows[0].Height);
        Assert.Equal(MergedCellValues.Restart, roundTripTable.Rows[0].Cells[0].HorizontalMerge);
        Assert.Equal(MergedCellValues.Continue, roundTripTable.Rows[0].Cells[1].HorizontalMerge);
        Assert.Equal(MergedCellValues.Restart, roundTripTable.Rows[0].Cells[0].VerticalMerge);
        Assert.Equal(MergedCellValues.Continue, roundTripTable.Rows[1].Cells[0].VerticalMerge);
        Assert.Equal(TableVerticalAlignmentValues.Center, roundTripTable.Rows[0].Cells[0].VerticalAlignment);
        Assert.Equal(TextDirectionValues.TopToBottomRightToLeftRotated, roundTripTable.Rows[0].Cells[0].TextDirection);
        Assert.Equal(TableWidthUnitValues.Dxa, roundTripTable.Rows[0].Cells[0].WidthType);
        Assert.Equal(1800, roundTripTable.Rows[0].Cells[0].Width);
        Assert.False(roundTripTable.Rows[0].Cells[0].WrapText);
        Assert.True(roundTripTable.Rows[0].Cells[0].FitText);
        Assert.Equal(TableVerticalAlignmentValues.Bottom, roundTripTable.Rows[1].Cells[2].VerticalAlignment);
        Assert.Equal("E6F2FF", roundTripTable.Rows[0].Cells[0].ShadingFillColorHex);
        Assert.Equal("00AA55", roundTripTable.Rows[0].Cells[0]._tableCellProperties?.Shading?.Color?.Value);
        Assert.Equal(ShadingPatternValues.Percent37, roundTripTable.Rows[0].Cells[0].ShadingPattern);
        Assert.Equal((short?)120, roundTripTable.Rows[0].Cells[0].MarginTopWidth);
        Assert.Equal((short?)180, roundTripTable.Rows[0].Cells[0].MarginLeftWidth);
        Assert.Equal((short?)240, roundTripTable.Rows[0].Cells[0].MarginBottomWidth);
        Assert.Equal((short?)300, roundTripTable.Rows[0].Cells[0].MarginRightWidth);
        Assert.Equal(BorderValues.Single, roundTripTable.Rows[0].Cells[0].Borders.TopStyle);
        Assert.Equal(12U, roundTripTable.Rows[0].Cells[0].Borders.TopSize?.Value);
        Assert.Equal("4472C4", roundTripTable.Rows[0].Cells[0].Borders.TopColorHex);
        Assert.Equal(BorderValues.Double, roundTripTable.Rows[0].Cells[0].Borders.LeftStyle);
        Assert.Equal(8U, roundTripTable.Rows[0].Cells[0].Borders.LeftSize?.Value);
        Assert.Equal("00AA55", roundTripTable.Rows[0].Cells[0].Borders.LeftColorHex);
        Assert.Equal(BorderValues.Dotted, roundTripTable.Rows[0].Cells[0].Borders.TopLeftToBottomRightStyle);
        Assert.Equal(6U, roundTripTable.Rows[0].Cells[0].Borders.TopLeftToBottomRightSize?.Value);
        Assert.Equal("4472C4", roundTripTable.Rows[0].Cells[0].Borders.TopLeftToBottomRightColorHex);
        Assert.Equal(BorderValues.Dashed, roundTripTable.Rows[0].Cells[0].Borders.TopRightToBottomLeftStyle);
        Assert.Equal(10U, roundTripTable.Rows[0].Cells[0].Borders.TopRightToBottomLeftSize?.Value);
        Assert.Equal("00AA55", roundTripTable.Rows[0].Cells[0].Borders.TopRightToBottomLeftColorHex);
    }
}
