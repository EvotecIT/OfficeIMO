using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfSeparatedTableTests {
    [Fact]
    public void BodyParagraphSeparatesNativeTablesWithoutSplittingConsecutiveRows() {
        const string rtf = @"{\rtf1\ansi
\trowd\cellx2000\pard\intbl Inner first\cell\row
\trowd\cellx2000\pard\intbl Inner second\cell\row
\pard Between planets\par
\trowd\cellx2000\pard\intbl Outer first\cell\row
\trowd\cellx2000\pard\intbl Outer second\cell\row
\pard After planets\par}";

        RtfDocument document = RtfDocument.Read(rtf).Document;

        Assert.Collection(document.Blocks,
            block => AssertRows(Assert.IsType<RtfTable>(block), "Inner first", "Inner second"),
            block => Assert.Equal("Between planets", Assert.IsType<RtfParagraph>(block).ToPlainText()),
            block => AssertRows(Assert.IsType<RtfTable>(block), "Outer first", "Outer second"),
            block => Assert.Equal("After planets", Assert.IsType<RtfParagraph>(block).ToPlainText()));
        RtfDocument reopened = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal(2, reopened.Blocks.OfType<RtfTable>().Count());
        Assert.Equal("Between planets", Assert.IsType<RtfParagraph>(reopened.Blocks[1]).ToPlainText());
    }

    private static void AssertRows(RtfTable table, params string[] expected) =>
        Assert.Equal(expected, table.Rows.Select(row => Assert.Single(Assert.Single(row.Cells).Paragraphs).ToPlainText()));
}
