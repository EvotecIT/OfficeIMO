using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlTableFormatTests {
    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Row_Direction() {
        const string html = "<table><tr dir=\"rtl\"><td>RTL</td></tr><tr style=\"direction:ltr\"><td>LTR</td></tr><tr dir=\"auto\"><td>Plain</td></tr></table>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

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

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(html.ToRtfDocumentFromHtml().Blocks));
        Assert.Equal(RtfTableRowDirection.RightToLeft, roundTripTable.Rows[0].Direction);
        Assert.Equal(RtfTableRowDirection.LeftToRight, roundTripTable.Rows[1].Direction);
    }

    [Fact]
    public void Html_ToRtfDocument_Parses_Table_Cell_Text_Flow() {
        const string html = "<table><tr><td style=\"writing-mode:vertical-rl\">Vertical</td><td style=\"writing-mode:sideways-lr\">Sideways</td><td style=\"--officeimo-rtf-text-flow:tb-rl-v\">Exact</td></tr></table>";

        RtfDocument document = html.ToRtfDocumentFromHtml();

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

        RtfTable roundTripTable = Assert.IsType<RtfTable>(Assert.Single(html.ToRtfDocumentFromHtml().Blocks));
        Assert.Equal(RtfTableCellTextFlow.TopToBottomRightToLeftVertical, roundTripTable.Rows[0].Cells[0].TextFlow);
        Assert.Equal(RtfTableCellTextFlow.LeftToRightTopToBottom, roundTripTable.Rows[0].Cells[1].TextFlow);
    }
}
