using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfLiteralStringParserTests {
    [Fact]
    public void PreservesPlainBytesAndPdfEscapes() {
        Assert.Equal(new byte[] { 0x41, 0x42, 0x43 }, PdfStringParser.ParseLiteralToBytes("ABC"));
        Assert.Equal(
            new byte[] { 0x41, 0x0A, 0x0D, 0x09, 0x08, 0x0C, 0x5C, 0x28, 0x29, 0x71 },
            PdfStringParser.ParseLiteralToBytes("A\\n\\r\\t\\b\\f\\\\\\(\\)\\q"));
        Assert.Equal(new byte[] { 0x41, 0x20, 0x07, 0x5A }, PdfStringParser.ParseLiteralToBytes("\\101\\40\\7Z"));
    }

    [Fact]
    public void DropsContinuedLinesAndTrailingEscape() {
        Assert.Empty(PdfStringParser.ParseLiteralToBytes(string.Empty));
        Assert.Equal(
            new byte[] { 0x41, 0x42, 0x43, 0x44 },
            PdfStringParser.ParseLiteralToBytes("A\\\r\nB\\\nC\\\rD\\"));
    }
}
