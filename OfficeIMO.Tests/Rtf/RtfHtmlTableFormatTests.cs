using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlTableFormatTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Row_Direction() {
        const string html = "<table><tr dir=\"rtl\"><td>RTL</td></tr><tr style=\"direction:ltr\"><td>LTR</td></tr><tr dir=\"auto\"><td>Plain</td></tr></table>";

        RtfDocument document = html.LoadRtfFromHtml();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        Assert.Equal(RtfTableRowDirection.RightToLeft, table.Rows[0].Direction);
        Assert.Equal(RtfTableRowDirection.LeftToRight, table.Rows[1].Direction);
        Assert.Null(table.Rows[2].Direction);

        string rtf = document.ToRtf();
        Assert.Contains(@"\rtlrow", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\ltrrow", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.Equal(RtfTableRowDirection.RightToLeft, roundTripTable.Rows[0].Direction);
        Assert.Equal(RtfTableRowDirection.LeftToRight, roundTripTable.Rows[1].Direction);
        Assert.Null(roundTripTable.Rows[2].Direction);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Row_Direction() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(2, 1);
        table.Rows[0]
            .SetDirection(RtfTableRowDirection.RightToLeft)
            .Cells[0]
            .AddParagraph("RTL");
        table.Rows[1]
            .SetDirection(RtfTableRowDirection.LeftToRight)
            .Cells[0]
            .AddParagraph("LTR");

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr dir=\"rtl\" style=\"direction:rtl;unicode-bidi:isolate;--officeimo-rtf-direction:rtl;\"><td><p>RTL</p></td></tr><tr dir=\"ltr\" style=\"direction:ltr;unicode-bidi:isolate;--officeimo-rtf-direction:ltr;\"><td><p>LTR</p></td></tr></tbody></table>", html);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(html.LoadRtfFromHtml().Blocks));
        Assert.Equal(RtfTableRowDirection.RightToLeft, roundTripTable.Rows[0].Direction);
        Assert.Equal(RtfTableRowDirection.LeftToRight, roundTripTable.Rows[1].Direction);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Cell_Text_Flow() {
        const string html = "<table><tr><td style=\"writing-mode:vertical-rl\">Vertical</td><td style=\"writing-mode:sideways-lr\">Sideways</td><td style=\"--officeimo-rtf-text-flow:tb-rl-v\">Exact</td></tr></table>";

        RtfDocument document = html.LoadRtfFromHtml();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeft, table.Rows[0].Cells[0].TextFlow);
        Assert.Equal(RtfTableCellTextFlow.BottomToTopLeftToRight, table.Rows[0].Cells[1].TextFlow);
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeftVertical, table.Rows[0].Cells[2].TextFlow);

        string rtf = document.ToRtf();
        Assert.Contains(@"\cltxtbrl", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cltxbtlr", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\cltxtbrlv", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeft, roundTripTable.Rows[0].Cells[0].TextFlow);
        Assert.Equal(RtfTableCellTextFlow.BottomToTopLeftToRight, roundTripTable.Rows[0].Cells[1].TextFlow);
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeftVertical, roundTripTable.Rows[0].Cells[2].TextFlow);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Cell_Text_Flow() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 2);
        table.Rows[0].Cells[0]
            .SetTextFlow(RtfTableCellTextFlow.TopToBottomRightToLeftVertical)
            .AddParagraph("Vertical");
        table.Rows[0].Cells[1]
            .SetTextFlow(RtfTableCellTextFlow.LeftToRightTopToBottom)
            .AddParagraph("Normal");

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr><td style=\"writing-mode:vertical-rl;text-orientation:upright;--officeimo-rtf-text-flow:tb-rl-v;\"><p>Vertical</p></td><td style=\"writing-mode:horizontal-tb;--officeimo-rtf-text-flow:ltr-tb;\"><p>Normal</p></td></tr></tbody></table>", html);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(html.LoadRtfFromHtml().Blocks));
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeftVertical, roundTripTable.Rows[0].Cells[0].TextFlow);
        Assert.Equal(RtfTableCellTextFlow.LeftToRightTopToBottom, roundTripTable.Rows[0].Cells[1].TextFlow);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Cell_Flags() {
        const string html = "<table><tr><td style=\"white-space:nowrap;--officeimo-rtf-hide-cell-mark:true;--officeimo-rtf-fit-text:true\">Flags</td><td style=\"--officeimo-rtf-cell-nowrap:true\">Custom</td></tr></table>";

        RtfDocument document = html.LoadRtfFromHtml();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        Assert.True(table.Rows[0].Cells[0].NoWrap);
        Assert.True(table.Rows[0].Cells[0].HideCellMark);
        Assert.True(table.Rows[0].Cells[0].FitText);
        Assert.True(table.Rows[0].Cells[1].NoWrap);

        string rtf = document.ToRtf();
        Assert.Contains(@"\clhidemark\clNoWrap\clFitText", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clNoWrap", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.True(roundTripTable.Rows[0].Cells[0].HideCellMark);
        Assert.True(roundTripTable.Rows[0].Cells[0].NoWrap);
        Assert.True(roundTripTable.Rows[0].Cells[0].FitText);
        Assert.True(roundTripTable.Rows[0].Cells[1].NoWrap);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Cell_Flags() {
        RtfDocument document = RtfDocument.Create();
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0]
            .SetHideCellMark()
            .SetNoWrap()
            .SetFitText()
            .AddParagraph("Flags");

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr><td style=\"white-space:nowrap;--officeimo-rtf-hide-cell-mark:true;--officeimo-rtf-cell-nowrap:true;--officeimo-rtf-fit-text:true;\"><p>Flags</p></td></tr></tbody></table>", html);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(html.LoadRtfFromHtml().Blocks));
        Assert.True(roundTripTable.Rows[0].Cells[0].HideCellMark);
        Assert.True(roundTripTable.Rows[0].Cells[0].NoWrap);
        Assert.True(roundTripTable.Rows[0].Cells[0].FitText);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Row_And_Cell_Shading_Metadata() {
        const string html = "<table><tr style=\"background-color:#eef6ff;--officeimo-rtf-shading-foreground:#00aa55;--officeimo-rtf-shading-pattern-value:7;--officeimo-rtf-shading-percent:6250;--officeimo-rtf-shading-pattern:dark-diagonal-cross\"><td style=\"background-color:#fff2cc;--officeimo-rtf-shading-foreground:#4472c4;--officeimo-rtf-shading-percent:37.5%;--officeimo-rtf-shading-pattern:clbgdkfdiag\">Cell</td></tr></table>";

        RtfDocument document = html.LoadRtfFromHtml();

        RtfTable table = Assert.IsType<RtfTable>(Assert.Single(document.Blocks));
        RtfTableRow row = table.Rows[0];
        RtfTableCell cell = row.Cells[0];
        AssertColor(document, row.BackgroundColorIndex, 0xEE, 0xF6, 0xFF);
        AssertColor(document, row.ShadingForegroundColorIndex, 0x00, 0xAA, 0x55);
        Assert.Equal(7, row.ShadingPatternValue);
        Assert.Equal(6250, row.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkDiagonalCross, row.ShadingPattern);
        AssertColor(document, cell.BackgroundColorIndex, 0xFF, 0xF2, 0xCC);
        AssertColor(document, cell.ShadingForegroundColorIndex, 0x44, 0x72, 0xC4);
        Assert.Equal(3750, cell.ShadingPatternPercent);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, cell.ShadingPattern);

        string rtf = document.ToRtf();
        Assert.Contains(@"\trcfpat", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trpat7", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trshdng6250", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\trbgdkdcross", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clcfpat", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clshdng3750", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\clbgdkfdiag", rtf, StringComparison.Ordinal);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(RtfDocument.Read(rtf).Document.Blocks));
        Assert.Equal(RtfShadingPattern.DarkDiagonalCross, roundTripTable.Rows[0].ShadingPattern);
        Assert.Equal(6250, roundTripTable.Rows[0].ShadingPatternPercent);
        Assert.Equal(7, roundTripTable.Rows[0].ShadingPatternValue);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, roundTripTable.Rows[0].Cells[0].ShadingPattern);
        Assert.Equal(3750, roundTripTable.Rows[0].Cells[0].ShadingPatternPercent);
    }

    [Fact]
    public void RtfDocument_ToHtml_Renders_Table_Row_And_Cell_Shading_Metadata() {
        RtfDocument document = RtfDocument.Create();
        int rowBackground = document.AddColor(0xEE, 0xF6, 0xFF);
        int rowForeground = document.AddColor(0x00, 0xAA, 0x55);
        int cellBackground = document.AddColor(0xFF, 0xF2, 0xCC);
        int cellForeground = document.AddColor(0x44, 0x72, 0xC4);
        RtfTable table = document.AddTable(1, 1);
        table.Rows[0].SetShading(rowBackground, rowForeground, patternValue: 7, patternPercent: 6250, pattern: RtfShadingPattern.DiagonalCross);
        table.Rows[0].Cells[0]
            .SetShading(cellBackground, cellForeground, patternPercent: 3750, pattern: RtfShadingPattern.DarkForwardDiagonal)
            .AddParagraph("Cell");

        string html = document.ToHtml();

        Assert.Equal("<table><tbody><tr style=\"background-color:#EEF6FF;--officeimo-rtf-shading-foreground:#00AA55;--officeimo-rtf-shading-pattern-value:7;--officeimo-rtf-shading-percent:6250;--officeimo-rtf-shading-pattern:diagonal-cross;\"><td style=\"background-color:#FFF2CC;--officeimo-rtf-shading-foreground:#4472C4;--officeimo-rtf-shading-percent:3750;--officeimo-rtf-shading-pattern:dark-forward-diagonal;\"><p>Cell</p></td></tr></tbody></table>", html);

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(html.LoadRtfFromHtml().Blocks));
        Assert.Equal(RtfShadingPattern.DiagonalCross, roundTripTable.Rows[0].ShadingPattern);
        Assert.Equal(6250, roundTripTable.Rows[0].ShadingPatternPercent);
        Assert.Equal(7, roundTripTable.Rows[0].ShadingPatternValue);
        Assert.Equal(RtfShadingPattern.DarkForwardDiagonal, roundTripTable.Rows[0].Cells[0].ShadingPattern);
        Assert.Equal(3750, roundTripTable.Rows[0].Cells[0].ShadingPatternPercent);
    }

    private static void AssertColor(RtfDocument document, int? colorIndex, byte red, byte green, byte blue) {
        Assert.True(colorIndex.HasValue);
        RtfColor color = document.Colors[colorIndex.Value - 1];
        Assert.Equal(red, color.Red);
        Assert.Equal(green, color.Green);
        Assert.Equal(blue, color.Blue);
    }
}
