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
}
