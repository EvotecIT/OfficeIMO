using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfSeparatedTableTests {
    [Theory]
    [InlineData("\\pard Between planets\\par", true)]
    [InlineData("\\sect", false)]
    public void BodyBoundarySeparatesNativeTablesWithoutSplittingConsecutiveRows(string separator, bool hasParagraph) {
        string rtf = @"{\rtf1\ansi
\trowd\cellx2000\pard\intbl Inner first\cell\row
\trowd\cellx2000\pard\intbl Inner second\cell\row" + separator + @"
\trowd\cellx2000\pard\intbl Outer first\cell\row
\trowd\cellx2000\pard\intbl Outer second\cell\row
\pard After planets\par}";

        RtfDocument document = RtfDocument.Read(rtf).Document;

        Assert.Equal(hasParagraph ? 4 : 3, document.Blocks.Count);
        AssertRows(Assert.IsType<RtfTable>(document.Blocks[0]), "Inner first", "Inner second");
        int secondTable = hasParagraph ? 2 : 1;
        if (hasParagraph) Assert.Equal("Between planets", Assert.IsType<RtfParagraph>(document.Blocks[1]).ToPlainText());
        AssertRows(Assert.IsType<RtfTable>(document.Blocks[secondTable]), "Outer first", "Outer second");
        Assert.Equal("After planets", Assert.IsType<RtfParagraph>(document.Blocks[secondTable + 1]).ToPlainText());
        RtfDocument reopened = RtfDocument.Read(document.ToRtf()).Document;
        Assert.Equal(2, reopened.Blocks.OfType<RtfTable>().Count());
        if (hasParagraph) Assert.Equal("Between planets", Assert.IsType<RtfParagraph>(reopened.Blocks[1]).ToPlainText());
    }

    private static void AssertRows(RtfTable table, params string[] expected) =>
        Assert.Equal(expected, table.Rows.Select(row => Assert.Single(Assert.Single(row.Cells).Paragraphs).ToPlainText()));
}
