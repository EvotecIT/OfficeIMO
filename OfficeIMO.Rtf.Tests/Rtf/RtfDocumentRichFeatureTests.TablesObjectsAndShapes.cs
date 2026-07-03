using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentRichFeatureTests {
    [Fact]
    public void Write_And_Read_Table_Block_With_Cell_Paragraphs() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].Cells[0].AddParagraph("A1");
        table.Rows[0].Cells[1].AddParagraph("B1").AddText(" bold").SetBold();
        table.Rows[1].Cells[0].AddParagraph("A2");
        table.Rows[1].Cells[1].AddParagraph("B2");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        RtfTable readTable = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.Equal(2, readTable.Rows.Count);
        Assert.Equal(2, readTable.Rows[0].Cells.Count);
        Assert.Equal("A1", Assert.Single(readTable.Rows[0].Cells[0].Paragraphs).ToPlainText());
        Assert.Equal("B1 bold", Assert.Single(readTable.Rows[0].Cells[1].Paragraphs).ToPlainText());
        Assert.Contains(readTable.Rows[0].Cells[1].Paragraphs[0].Runs, run => run.Text == " bold" && run.Bold);
        Assert.Contains(@"\trowd", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cellx2400", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\row", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Write_And_Read_Table_Row_And_Cell_Formatting() {
        RtfDocument document = RtfDocument.Create();
        int shadeColor = document.AddColor(0xEE, 0xF6, 0xFF);
        int borderColor = document.AddColor(0x44, 0x72, 0xC4);
        int patternColor = document.AddColor(0x00, 0xAA, 0x55);
        RtfTable table = document.AddTable(1, 3);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].KeepTogether = true;
        table.Rows[0].KeepWithNext = true;
        table.Rows[0].AutoFit = true;
        table.Rows[0].Direction = RtfTableRowDirection.RightToLeft;
        table.Rows[0].HeightTwips = 520;
        table.Rows[0].CellGapTwips = 240;
        table.Rows[0].LeftIndentTwips = 720;
        table.Rows[0].Alignment = RtfTableAlignment.Center;
        table.Rows[0].PreferredWidthUnit = RtfTableWidthUnit.Percent;
        table.Rows[0].PreferredWidth = 5000;
        table.Rows[0].SetShading(shadeColor, foregroundColorIndex: patternColor, patternValue: 7, patternPercent: 6250, pattern: RtfShadingPattern.DiagonalCross);
        table.Rows[0].SetPadding(topTwips: 120, leftTwips: 180, bottomTwips: 240, rightTwips: 300);
        table.Rows[0].SetSpacing(topTwips: 20, leftTwips: 30, bottomTwips: 40, rightTwips: 50);
        table.Rows[0].NoOverlap = true;
        table.Rows[0].SetPositionAnchors(RtfTableHorizontalAnchor.Margin, RtfTableVerticalAnchor.Page);
        table.Rows[0].SetPosition(RtfTableHorizontalPosition.Center, verticalPosition: RtfTableVerticalPosition.Bottom);
        table.Rows[0].SetTextWrapDistances(leftTwips: 187, rightTwips: 188, topTwips: 189, bottomTwips: 190);
        table.Rows[0].TopBorder.Style = RtfTableCellBorderStyle.Single;
        table.Rows[0].TopBorder.Width = 16;
        table.Rows[0].TopBorder.ColorIndex = borderColor;
        table.Rows[0].LeftBorder.Style = RtfTableCellBorderStyle.Double;
        table.Rows[0].LeftBorder.Width = 10;
        table.Rows[0].BottomBorder.Style = RtfTableCellBorderStyle.Dotted;
        table.Rows[0].RightBorder.Style = RtfTableCellBorderStyle.Dashed;
        table.Rows[0].HorizontalBorder.Style = RtfTableCellBorderStyle.Single;
        table.Rows[0].VerticalBorder.Style = RtfTableCellBorderStyle.Double;
        table.Rows[0].Cells[0].HorizontalMerge = RtfTableCellMerge.First;
        table.Rows[0].Cells[0].VerticalMerge = RtfTableCellMerge.First;
        table.Rows[0].Cells[0].SetShading(shadeColor, foregroundColorIndex: patternColor, patternPercent: 3750, pattern: RtfShadingPattern.DarkForwardDiagonal);
        table.Rows[0].Cells[0].VerticalAlignment = RtfTableCellVerticalAlignment.Center;
        table.Rows[0].Cells[0].TextFlow = RtfTableCellTextFlow.TopToBottomRightToLeftVertical;
        table.Rows[0].Cells[0].SetPreferredWidth(1800, RtfTableWidthUnit.Twips);
        table.Rows[0].Cells[0].SetHideCellMark();
        table.Rows[0].Cells[0].SetNoWrap();
        table.Rows[0].Cells[0].SetFitText();
        table.Rows[0].Cells[0].SetPadding(topTwips: 120, leftTwips: 180, bottomTwips: 240, rightTwips: 300);
        table.Rows[0].Cells[0].TopBorder.Style = RtfTableCellBorderStyle.Single;
        table.Rows[0].Cells[0].TopBorder.Width = 12;
        table.Rows[0].Cells[0].TopBorder.ColorIndex = borderColor;
        table.Rows[0].Cells[0].LeftBorder.Style = RtfTableCellBorderStyle.Double;
        table.Rows[0].Cells[0].LeftBorder.Width = 8;
        table.Rows[0].Cells[0].BottomBorder.Style = RtfTableCellBorderStyle.Dotted;
        table.Rows[0].Cells[0].RightBorder.Style = RtfTableCellBorderStyle.Dashed;
        table.Rows[0].Cells[0].TopLeftToBottomRightBorder.Style = RtfTableCellBorderStyle.Dotted;
        table.Rows[0].Cells[0].TopLeftToBottomRightBorder.Width = 6;
        table.Rows[0].Cells[0].TopLeftToBottomRightBorder.ColorIndex = borderColor;
        table.Rows[0].Cells[0].TopRightToBottomLeftBorder.Style = RtfTableCellBorderStyle.Dashed;
        table.Rows[0].Cells[0].TopRightToBottomLeftBorder.Width = 10;
        table.Rows[0].Cells[0].AddParagraph("Merged");
        table.Rows[0].Cells[1].HorizontalMerge = RtfTableCellMerge.Continue;
        table.Rows[0].Cells[1].AddParagraph("Continue");
        table.Rows[0].Cells[2].VerticalAlignment = RtfTableCellVerticalAlignment.Bottom;
        table.Rows[0].Cells[2].AddParagraph("Bottom");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        RtfTable readTable = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableRow readRow = Assert.Single(readTable.Rows);
        Assert.True(readRow.RepeatHeader);
        Assert.True(readRow.KeepTogether);
        Assert.True(readRow.KeepWithNext);
        Assert.True(readRow.AutoFit);
        Assert.Equal(RtfTableRowDirection.RightToLeft, readRow.Direction);
        Assert.Equal(520, readRow.HeightTwips);
        Assert.Equal(240, readRow.CellGapTwips);
        Assert.Equal(720, readRow.LeftIndentTwips);
        Assert.Equal(RtfTableAlignment.Center, readRow.Alignment);
        Assert.Equal(RtfTableWidthUnit.Percent, readRow.PreferredWidthUnit);
        Assert.Equal(5000, readRow.PreferredWidth);
        Assert.Equal(shadeColor, readRow.BackgroundColorIndex);
        Assert.Equal(patternColor, readRow.ShadingForegroundColorIndex);
        Assert.Equal(7, readRow.ShadingPatternValue);
        Assert.Equal(6250, readRow.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DiagonalCross, readRow.ShadingPattern);
        Assert.Equal(120, readRow.PaddingTopTwips);
        Assert.Equal(180, readRow.PaddingLeftTwips);
        Assert.Equal(240, readRow.PaddingBottomTwips);
        Assert.Equal(300, readRow.PaddingRightTwips);
        Assert.Equal(20, readRow.SpacingTopTwips);
        Assert.Equal(30, readRow.SpacingLeftTwips);
        Assert.Equal(40, readRow.SpacingBottomTwips);
        Assert.Equal(50, readRow.SpacingRightTwips);
        Assert.True(readRow.NoOverlap);
        Assert.Equal(RtfTableHorizontalAnchor.Margin, readRow.HorizontalAnchor);
        Assert.Equal(RtfTableVerticalAnchor.Page, readRow.VerticalAnchor);
        Assert.Equal(RtfTableHorizontalPosition.Center, readRow.HorizontalPosition);
        Assert.Null(readRow.HorizontalPositionTwips);
        Assert.Equal(RtfTableVerticalPosition.Bottom, readRow.VerticalPosition);
        Assert.Null(readRow.VerticalPositionTwips);
        Assert.Equal(187, readRow.TextWrapLeftTwips);
        Assert.Equal(188, readRow.TextWrapRightTwips);
        Assert.Equal(189, readRow.TextWrapTopTwips);
        Assert.Equal(190, readRow.TextWrapBottomTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, readRow.TopBorder.Style);
        Assert.Equal(16, readRow.TopBorder.Width);
        Assert.Equal(borderColor, readRow.TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Double, readRow.LeftBorder.Style);
        Assert.Equal(10, readRow.LeftBorder.Width);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, readRow.BottomBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, readRow.RightBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Single, readRow.HorizontalBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Double, readRow.VerticalBorder.Style);
        Assert.Equal(RtfTableCellMerge.First, readRow.Cells[0].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.Continue, readRow.Cells[1].HorizontalMerge);
        Assert.Equal(RtfTableCellMerge.First, readRow.Cells[0].VerticalMerge);
        Assert.Equal(shadeColor, readRow.Cells[0].BackgroundColorIndex);
        Assert.Equal(patternColor, readRow.Cells[0].ShadingForegroundColorIndex);
        Assert.Equal(3750, readRow.Cells[0].ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, readRow.Cells[0].ShadingPattern);
        Assert.Equal(RtfTableCellVerticalAlignment.Center, readRow.Cells[0].VerticalAlignment);
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeftVertical, readRow.Cells[0].TextFlow);
        Assert.Equal(RtfTableWidthUnit.Twips, readRow.Cells[0].PreferredWidthUnit);
        Assert.Equal(1800, readRow.Cells[0].PreferredWidth);
        Assert.True(readRow.Cells[0].HideCellMark);
        Assert.True(readRow.Cells[0].NoWrap);
        Assert.True(readRow.Cells[0].FitText);
        Assert.Equal(RtfTableCellVerticalAlignment.Bottom, readRow.Cells[2].VerticalAlignment);
        Assert.Equal(120, readRow.Cells[0].PaddingTopTwips);
        Assert.Equal(180, readRow.Cells[0].PaddingLeftTwips);
        Assert.Equal(240, readRow.Cells[0].PaddingBottomTwips);
        Assert.Equal(300, readRow.Cells[0].PaddingRightTwips);
        Assert.Equal(RtfTableCellBorderStyle.Single, readRow.Cells[0].TopBorder.Style);
        Assert.Equal(12, readRow.Cells[0].TopBorder.Width);
        Assert.Equal(borderColor, readRow.Cells[0].TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Double, readRow.Cells[0].LeftBorder.Style);
        Assert.Equal(8, readRow.Cells[0].LeftBorder.Width);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, readRow.Cells[0].BottomBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, readRow.Cells[0].RightBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, readRow.Cells[0].TopLeftToBottomRightBorder.Style);
        Assert.Equal(6, readRow.Cells[0].TopLeftToBottomRightBorder.Width);
        Assert.Equal(borderColor, readRow.Cells[0].TopLeftToBottomRightBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, readRow.Cells[0].TopRightToBottomLeftBorder.Style);
        Assert.Equal(10, readRow.Cells[0].TopRightToBottomLeftBorder.Width);
        Assert.Contains(@"\trhdr", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trkeep", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trkeepfollow", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trautofit1\rtlrow", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trrh520", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trgaph240", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trleft720", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trftsWidth2\trwWidth5000", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trcbpat1\trcfpat3\trpat7\trshdng6250\trbgdcross", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trpaddt120\trpaddft3\trpaddl180\trpaddfl3\trpaddb240\trpaddfb3\trpaddr300\trpaddfr3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trspdt20\trspdft3\trspdl30\trspdfl3\trspdb40\trspdfb3\trspdr50\trspdfr3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\tabsnoovrlp\tphmrg\tpvpg\tposxc\tposyb\tdfrmtxtLeft187\tdfrmtxtRight188\tdfrmtxtTop189\tdfrmtxtBottom190", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trqc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrt\brdrs\brdrw16\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrl\brdrdb\brdrw10", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrb\brdrdot", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrr\brdrdash", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrh\brdrs", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbrdrv\brdrdb", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clmgf", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clmrg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvmgf", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clcbpat1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clcfpat3\clshdng3750\clbgdkfdiag", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvertalc", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cltxtbrlv", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clftsWidth3\clwWidth1800\clhidemark\clNoWrap\clFitText", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clvertalb", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadt120\clpadft3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadl180\clpadfl3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadb240\clpadfb3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clpadr300\clpadfr3", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrt\brdrs\brdrw12\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrl\brdrdb\brdrw8", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrb\brdrdot", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbrdrr\brdrdash", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cldglu\brdrdot\brdrw6\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cldgll\brdrdash\brdrw10", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Read_Binds_Table_Row_Shading_Controls() {
        const string rtf = @"{\rtf1\ansi\trowd\trcbpat1\trcfpat2\trpat5\trshdng6250\trbgdkhor\cellx2400\pard\intbl A\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableRow row = Assert.Single(table.Rows);
        Assert.Equal(1, row.BackgroundColorIndex);
        Assert.Equal(2, row.ShadingForegroundColorIndex);
        Assert.Equal(5, row.ShadingPatternValue);
        Assert.Equal(6250, row.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, row.ShadingPattern);
    }

    [Fact]
    public void Read_Binds_Table_Row_Padding_And_Spacing_Controls() {
        const string rtf = @"{\rtf1\ansi\trowd\trpaddft3\trpaddt120\trpaddl180\trpaddfl3\trpaddfb3\trpaddb240\trpaddr300\trpaddfr3\trspdft3\trspdt20\trspdl30\trspdfl3\trspdfb3\trspdb40\trspdr50\trspdfr3\cellx2400\pard\intbl A\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableRow row = Assert.Single(table.Rows);
        Assert.Equal(120, row.PaddingTopTwips);
        Assert.Equal(180, row.PaddingLeftTwips);
        Assert.Equal(240, row.PaddingBottomTwips);
        Assert.Equal(300, row.PaddingRightTwips);
        Assert.Equal(20, row.SpacingTopTwips);
        Assert.Equal(30, row.SpacingLeftTwips);
        Assert.Equal(40, row.SpacingBottomTwips);
        Assert.Equal(50, row.SpacingRightTwips);
    }

    [Fact]
    public void Read_Binds_Table_Row_AutoFit_And_Direction_Controls() {
        const string rtf = @"{\rtf1\ansi\trowd\trautofit0\ltrrow\cellx2400\pard\intbl A\cell\row\trowd\trautofit1\rtlrow\cellx2400\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.False(table.Rows[0].AutoFit);
        Assert.Equal(RtfTableRowDirection.LeftToRight, table.Rows[0].Direction);
        Assert.True(table.Rows[1].AutoFit);
        Assert.Equal(RtfTableRowDirection.RightToLeft, table.Rows[1].Direction);
    }

    [Fact]
    public void Write_Emits_Explicit_Table_Row_AutoFit_Off() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].SetAutoFit(false).SetDirection(RtfTableRowDirection.LeftToRight);
        table.Rows[0].Cells[0].AddParagraph("Fixed");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\trautofit0\ltrrow", rtf, StringComparison.Ordinal);
        RtfTable readTable = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.False(readTable.Rows[0].AutoFit);
        Assert.Equal(RtfTableRowDirection.LeftToRight, readTable.Rows[0].Direction);
    }

    [Fact]
    public void Read_Binds_Positioned_Table_Row_Controls() {
        const string rtf = @"{\rtf1\ansi\trowd\tabsnoovrlp\tphmrg\tpvpara\tposnegx-120\tposnegy-240\tdfrmtxtLeft187\tdfrmtxtRight188\tdfrmtxtTop189\tdfrmtxtBottom190\cellx2400\pard\intbl A\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableRow row = Assert.Single(table.Rows);
        Assert.True(row.NoOverlap);
        Assert.Equal(RtfTableHorizontalAnchor.Margin, row.HorizontalAnchor);
        Assert.Equal(RtfTableVerticalAnchor.Paragraph, row.VerticalAnchor);
        Assert.Equal(RtfTableHorizontalPosition.NegativeAbsolute, row.HorizontalPosition);
        Assert.Equal(-120, row.HorizontalPositionTwips);
        Assert.Equal(RtfTableVerticalPosition.NegativeAbsolute, row.VerticalPosition);
        Assert.Equal(-240, row.VerticalPositionTwips);
        Assert.Equal(187, row.TextWrapLeftTwips);
        Assert.Equal(188, row.TextWrapRightTwips);
        Assert.Equal(189, row.TextWrapTopTwips);
        Assert.Equal(190, row.TextWrapBottomTwips);
    }

    [Fact]
    public void Write_Emits_Positioned_Table_Row_Absolute_Controls() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].NoOverlap = true;
        table.Rows[0].SetPositionAnchors(RtfTableHorizontalAnchor.Page, RtfTableVerticalAnchor.Margin);
        table.Rows[0].SetPosition(RtfTableHorizontalPosition.Absolute, horizontalTwips: 720, verticalPosition: RtfTableVerticalPosition.Absolute, verticalTwips: 1440);
        table.Rows[0].SetTextWrapDistances(leftTwips: 10, rightTwips: 20, topTwips: 30, bottomTwips: 40);
        table.Rows[0].Cells[0].AddParagraph("Positioned");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"\tabsnoovrlp\tphpg\tpvmrg\tposx720\tposy1440\tdfrmtxtLeft10\tdfrmtxtRight20\tdfrmtxtTop30\tdfrmtxtBottom40", rtf, StringComparison.Ordinal);
        RtfTable readTable = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.Equal(RtfTableHorizontalAnchor.Page, readTable.Rows[0].HorizontalAnchor);
        Assert.Equal(RtfTableVerticalAnchor.Margin, readTable.Rows[0].VerticalAnchor);
        Assert.Equal(RtfTableHorizontalPosition.Absolute, readTable.Rows[0].HorizontalPosition);
        Assert.Equal(720, readTable.Rows[0].HorizontalPositionTwips);
        Assert.Equal(RtfTableVerticalPosition.Absolute, readTable.Rows[0].VerticalPosition);
        Assert.Equal(1440, readTable.Rows[0].VerticalPositionTwips);
    }

    [Fact]
    public void Read_Binds_Table_Cell_Canonical_Dark_Horizontal_Shading_Control() {
        const string rtf = @"{\rtf1\ansi\trowd\clcbpat1\clcfpat2\clshdng3750\clbgdkhor\cellx2400\pard\intbl A\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableCell cell = Assert.Single(Assert.Single(table.Rows).Cells);
        Assert.Equal(1, cell.BackgroundColorIndex);
        Assert.Equal(2, cell.ShadingForegroundColorIndex);
        Assert.Equal(3750, cell.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkHorizontal, cell.ShadingPattern);
    }

    [Fact]
    public void Read_Binds_Table_Cell_Text_Flow_Controls() {
        const string rtf = @"{\rtf1\ansi\trowd\cltxlrtb\cellx1200\cltxtbrl\cellx2400\cltxbtlr\cellx3600\cltxlrtbv\cellx4800\cltxtbrlv\cellx6000\pard\intbl A\cell\pard\intbl B\cell\pard\intbl C\cell\pard\intbl D\cell\pard\intbl E\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.Collection(table.Rows[0].Cells,
            cell => Assert.Equal(RtfTableCellTextFlow.LeftToRightTopToBottom, cell.TextFlow),
            cell => Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeft, cell.TextFlow),
            cell => Assert.Equal(RtfTableCellTextFlow.BottomToTopLeftToRight, cell.TextFlow),
            cell => Assert.Equal(RtfTableCellTextFlow.LeftToRightTopToBottomVertical, cell.TextFlow),
            cell => Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeftVertical, cell.TextFlow));
    }

    [Fact]
    public void Read_Binds_Table_Cell_Preferred_Widths_Per_Cell() {
        const string rtf = @"{\rtf1\ansi\trowd\clftsWidth2\clwWidth2500\cellx2400\clftsWidth3\clwWidth1440\clhidemark\cellx4800\pard\intbl A\cell\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.Equal(RtfTableWidthUnit.Percent, table.Rows[0].Cells[0].PreferredWidthUnit);
        Assert.Equal(2500, table.Rows[0].Cells[0].PreferredWidth);
        Assert.False(table.Rows[0].Cells[0].HideCellMark);
        Assert.Equal(RtfTableWidthUnit.Twips, table.Rows[0].Cells[1].PreferredWidthUnit);
        Assert.Equal(1440, table.Rows[0].Cells[1].PreferredWidth);
        Assert.True(table.Rows[0].Cells[1].HideCellMark);
    }

    [Fact]
    public void Read_Binds_Table_Cell_NoWrap_And_FitText() {
        const string rtf = @"{\rtf1\ansi\trowd\clNoWrap\clFitText\cellx2400\cellx4800\pard\intbl A\cell\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.True(table.Rows[0].Cells[0].NoWrap);
        Assert.True(table.Rows[0].Cells[0].FitText);
        Assert.False(table.Rows[0].Cells[1].NoWrap);
        Assert.False(table.Rows[0].Cells[1].FitText);
    }

    [Fact]
    public void Read_Binds_Table_Row_Positioning_Per_Row() {
        const string rtf = @"{\rtf1\ansi\trowd\trleft360\trql\cellx2400\pard\intbl A\cell\row\trowd\trleft720\trqr\cellx2400\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.Equal(360, table.Rows[0].LeftIndentTwips);
        Assert.Equal(RtfTableAlignment.Left, table.Rows[0].Alignment);
        Assert.Equal(720, table.Rows[1].LeftIndentTwips);
        Assert.Equal(RtfTableAlignment.Right, table.Rows[1].Alignment);
    }

    [Fact]
    public void Read_Binds_Table_Row_Cell_Gap_Per_Row() {
        const string rtf = @"{\rtf1\ansi\trowd\trgaph120\cellx2400\pard\intbl A\cell\row\trowd\trgaph360\cellx2400\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.Equal(120, table.Rows[0].CellGapTwips);
        Assert.Equal(360, table.Rows[1].CellGapTwips);
        Assert.Equal("A", Assert.Single(table.Rows[0].Cells[0].Paragraphs).ToPlainText());
        Assert.Equal("B", Assert.Single(table.Rows[1].Cells[0].Paragraphs).ToPlainText());
    }

    [Fact]
    public void Read_Binds_Table_Row_Keep_Controls() {
        const string rtf = @"{\rtf1\ansi\trowd\trkeep\trkeepfollow\cellx2400\pard\intbl A\cell\row\trowd\cellx2400\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.True(table.Rows[0].KeepTogether);
        Assert.True(table.Rows[0].KeepWithNext);
        Assert.False(table.Rows[1].KeepTogether);
        Assert.False(table.Rows[1].KeepWithNext);
    }

    [Fact]
    public void Read_Binds_Table_Preferred_Width_Per_Row() {
        const string rtf = @"{\rtf1\ansi\trowd\trftsWidth2\trwWidth4250\cellx2400\pard\intbl A\cell\row\trowd\trftsWidth3\trwWidth7200\cellx2400\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        Assert.Equal(RtfTableWidthUnit.Percent, table.Rows[0].PreferredWidthUnit);
        Assert.Equal(4250, table.Rows[0].PreferredWidth);
        Assert.Equal(RtfTableWidthUnit.Twips, table.Rows[1].PreferredWidthUnit);
        Assert.Equal(7200, table.Rows[1].PreferredWidth);
    }

    [Fact]
    public void Read_Binds_Table_Row_Borders() {
        const string rtf = @"{\rtf1\ansi\trowd\trbrdrt\brdrs\brdrw12\brdrcf1\trbrdrl\brdrdb\brdrw8\trbrdrb\brdrdot\trbrdrr\brdrdash\trbrdrh\brdrs\trbrdrv\brdrdb\cellx2400\pard\intbl A\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableRow row = Assert.Single(table.Rows);
        Assert.Equal(RtfTableCellBorderStyle.Single, row.TopBorder.Style);
        Assert.Equal(12, row.TopBorder.Width);
        Assert.Equal(1, row.TopBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Double, row.LeftBorder.Style);
        Assert.Equal(8, row.LeftBorder.Width);
        Assert.Equal(RtfTableCellBorderStyle.Dotted, row.BottomBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Dashed, row.RightBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Single, row.HorizontalBorder.Style);
        Assert.Equal(RtfTableCellBorderStyle.Double, row.VerticalBorder.Style);
    }

    [Fact]
    public void Read_Binds_Table_Cell_Diagonal_Borders() {
        const string rtf = @"{\rtf1\ansi\trowd\cldglu\brdrs\brdrw12\brdrcf1\cldgll\brdrdb\brdrw8\cellx2400\pard\intbl A\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableCell cell = table.Rows[0].Cells[0];
        Assert.Equal(RtfTableCellBorderStyle.Single, cell.TopLeftToBottomRightBorder.Style);
        Assert.Equal(12, cell.TopLeftToBottomRightBorder.Width);
        Assert.Equal(1, cell.TopLeftToBottomRightBorder.ColorIndex);
        Assert.Equal(RtfTableCellBorderStyle.Double, cell.TopRightToBottomLeftBorder.Style);
        Assert.Equal(8, cell.TopRightToBottomLeftBorder.Width);
    }

    [Fact]
    public void Read_Binds_Table_Cell_Padding_When_Units_Are_Twips() {
        const string rtf = @"{\rtf1\ansi\trowd\clpadt120\clpadft3\clpadl180\clpadfl3\clpadb240\clpadfb3\clpadr300\clpadfr3\cellx2400\clpadt20\clpadft2\cellx4800\pard\intbl A\cell\pard\intbl B\cell\row}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(read.Document.Blocks));
        RtfTableCell padded = table.Rows[0].Cells[0];
        Assert.Equal(120, padded.PaddingTopTwips);
        Assert.Equal(180, padded.PaddingLeftTwips);
        Assert.Equal(240, padded.PaddingBottomTwips);
        Assert.Equal(300, padded.PaddingRightTwips);
        Assert.Null(table.Rows[0].Cells[1].PaddingTopTwips);
    }

    [Fact]
    public void Write_And_Read_Png_Picture_Block_Without_Image_Dependencies() {
        byte[] pngSignature = { 0x89, 0x50, 0x4E, 0x47 };
        RtfDocument document = RtfDocument.Create();
        RtfImage image = document.AddImage(RtfImageFormat.Png, pngSignature);
        image.SourceWidth = 10;
        image.SourceHeight = 20;
        image.DesiredWidthTwips = 1440;
        image.DesiredHeightTwips = 2880;

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        RtfImage readImage = Assert.IsType<RtfImage>(Assert.Single(read.Document.Blocks));
        Assert.Equal(RtfImageFormat.Png, readImage.Format);
        Assert.Equal(pngSignature, readImage.Data);
        Assert.Equal(10, readImage.SourceWidth);
        Assert.Equal(20, readImage.SourceHeight);
        Assert.Equal(1440, readImage.DesiredWidthTwips);
        Assert.Equal(2880, readImage.DesiredHeightTwips);
        Assert.Contains(@"{\pict\pngblip\picw10\pich20\picwgoal1440\pichgoal2880", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Write_And_Read_Png_Picture_Inline_In_Paragraph_Order() {
        byte[] pngSignature = { 0x89, 0x50, 0x4E, 0x47 };
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before ");
        RtfImage image = paragraph.AddImage(RtfImageFormat.Png, pngSignature);
        image.SourceWidth = 10;
        image.SourceHeight = 20;
        image.DesiredWidthTwips = 1440;
        image.DesiredHeightTwips = 2880;
        paragraph.AddText(" after");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Before  after", readParagraph.ToPlainText());
        Assert.Collection(readParagraph.Inlines,
            inline => Assert.Equal("Before ", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfImage readImage = Assert.IsType<RtfImage>(inline);
                Assert.Equal(RtfImageFormat.Png, readImage.Format);
                Assert.Equal(pngSignature, readImage.Data);
                Assert.Equal(10, readImage.SourceWidth);
                Assert.Equal(20, readImage.SourceHeight);
                Assert.Equal(1440, readImage.DesiredWidthTwips);
                Assert.Equal(2880, readImage.DesiredHeightTwips);
            },
            inline => Assert.Equal(" after", Assert.IsType<RtfRun>(inline).Text));
        Assert.Contains(@"{\pict\pngblip\picw10\pich20\picwgoal1440\pichgoal2880", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Write_And_Read_Embedded_Object_Without_Ole_Dependencies() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before ");
        RtfObject rtfObject = paragraph.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2, 3, 255 });
        rtfObject.ClassName = "Package";
        rtfObject.Name = "Attachment";
        rtfObject.Width = 100;
        rtfObject.Height = 200;
        rtfObject.ScaleX = 75;
        rtfObject.ScaleY = 80;
        rtfObject.Result.AddText("Display").SetBold();
        paragraph.AddText(" after");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\object\objemb\objw100\objh200\objscalex75\objscaley80{\*\objclass Package}{\*\objname Attachment}{\*\objdata 010203ff}{\result \b Display\b0 }}", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Before Display after", readParagraph.ToPlainText());
        RtfObject readObject = Assert.IsType<RtfObject>(readParagraph.Inlines[1]);
        Assert.Equal(RtfObjectKind.Embedded, readObject.Kind);
        Assert.Equal("Package", readObject.ClassName);
        Assert.Equal("Attachment", readObject.Name);
        Assert.Equal(new byte[] { 1, 2, 3, 255 }, readObject.Data);
        Assert.Equal(100, readObject.Width);
        Assert.Equal(200, readObject.Height);
        Assert.Equal(75, readObject.ScaleX);
        Assert.Equal(80, readObject.ScaleY);
        Assert.Equal("Display", readObject.Result.ToPlainText());
        Assert.Contains(readObject.Result.Runs, run => run.Text == "Display" && run.Bold);
    }

    [Fact]
    public void Read_Binds_Shape_TextBox_Instructions_And_Properties_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\shp{\*\shpinst\shpleft100\shptop200\shpright2100\shpbottom900{\sp{\sn shapeType}{\sv 202}}{\sp{\sn fLine}{\sv 0}}{\shptxt\pard Text box\par}}}\pard Body\par}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        Assert.Collection(read.Document.Blocks,
            block => {
                RtfShape shape = Assert.IsType<RtfShape>(block);
                Assert.Contains(shape.Instructions, instruction => instruction.Name == "shpleft" && instruction.Parameter == 100);
                Assert.Contains(shape.Instructions, instruction => instruction.Name == "shptop" && instruction.Parameter == 200);
                Assert.Contains(shape.Instructions, instruction => instruction.Name == "shpright" && instruction.Parameter == 2100);
                Assert.Contains(shape.Instructions, instruction => instruction.Name == "shpbottom" && instruction.Parameter == 900);
                Assert.Contains(shape.Properties, property => property.Name == "shapeType" && property.Value == "202");
                Assert.Contains(shape.Properties, property => property.Name == "fLine" && property.Value == "0");
                RtfParagraph textBoxParagraph = Assert.Single(shape.TextBoxParagraphs);
                Assert.Equal("Text box", textBoxParagraph.ToPlainText());
            },
            block => Assert.Equal("Body", Assert.IsType<RtfParagraph>(block).ToPlainText()));
    }

    [Fact]
    public void Write_And_Read_Shape_TextBox_Without_Drawing_Dependencies() {
        RtfDocument document = RtfDocument.Create();
        RtfShape shape = document.AddShape();
        shape.AddInstruction("shpleft", 100);
        shape.AddInstruction("shptop", 200);
        shape.AddInstruction("shpright", 2100);
        shape.AddInstruction("shpbottom", 900);
        shape.AddProperty("shapeType", "202");
        shape.AddProperty("fLine", "0");
        shape.AddTextBoxParagraph("Text box");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\shp{\*\shpinst\shpleft100\shptop200\shpright2100\shpbottom900", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\sp{\sn shapeType}{\sv 202}}{\sp{\sn fLine}{\sv 0}}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\shptxt\pard\ql Text box\par", rtf, StringComparison.Ordinal);
        RtfShape roundTrip = Assert.IsType<RtfShape>(Assert.Single(read.Document.Blocks));
        Assert.Equal("Text box", roundTrip.ToPlainText());
        Assert.Contains(roundTrip.Properties, property => property.Name == "shapeType" && property.Value == "202");
        Assert.Contains(roundTrip.Instructions, instruction => instruction.Name == "shpbottom" && instruction.Parameter == 900);
    }
}
