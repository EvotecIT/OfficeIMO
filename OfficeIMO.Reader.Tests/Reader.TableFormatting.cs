using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderTableFormattingTests {
    [Fact]
    public void ExportRetainsMarkupAndBackslashesWhileNormalizingCellLineBreaks() {
        var table = new ReaderTable {
            Columns = new[] { "Value" },
            Rows = new[] { new[] { "*bold* | <b>x</b>\r\nnext\\path\rthird\nfourth" } }
        };
        Assert.Equal("| Value |" + Environment.NewLine + "| --- |" + Environment.NewLine
            + "| *bold* \\| <b>x</b><br>next\\path<br>third<br>fourth |", table.ToMarkdownTable());
    }

    [Fact]
    public void CsvProjectionRetainsQuotedCellLineBreaksAndItsBackslashEscaping() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("Value\r\n\"*bold* | <b>x</b>\r\nnext\\path\"\r\n"));
        ReaderChunk chunk = Assert.Single(CsvReaderAdapter.Read(stream, "values.csv",
            csvOptions: new CsvReadOptions { IncludeMarkdown = true }));
        Assert.Equal("| Value |" + Environment.NewLine + "| --- |" + Environment.NewLine
            + "| *bold* \\| <b>x</b>\r\nnext\\\\path |", chunk.Markdown);
    }
}
