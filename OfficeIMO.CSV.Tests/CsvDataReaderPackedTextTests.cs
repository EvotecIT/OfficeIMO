using System;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderPackedTextTests {
    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    public void BufferedReadersRetainEmptyFieldsAfterQuotedPrefix(int transport, bool trailing) {
        string first = new string('a', 40);
        string last = trailing ? "" : new string('x', 40);
        string rowText = first + ",b,c,d,\"q\",,," + last + "\r\n";
        string csv = "A,B,C,D,E,F,G,H\r\n" + rowText + rowText;
        using var bytes = new MemoryStream(Encoding.UTF8.GetBytes(csv));
        using TextReader text = transport == 2 ? new StringReader(csv) : new StreamReader(bytes);
        using var reader = transport == 1
            ? CsvDocument.OpenReader(bytes, new CsvReaderOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict })
            : CsvDocument.OpenReader(text, new CsvReaderOptions { ColumnCountMismatchPolicy = CsvColumnCountMismatchPolicy.Strict });
        string[] expected = { first, "b", "c", "d", "q", "", "", last };
        for (int row = 0; row < 2; row++) {
            Assert.True(reader.Read());
            for (int column = 0; column < expected.Length; column++) Assert.Equal(expected[column], reader.GetString(column));
        }
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SnapshotStrictQuotesRejectWhitespaceBeforeOpeningQuote(bool oversizedPrefix) {
        string padding = oversizedPrefix ? new string('x', 70_000) : "";
        using var reader = CsvDocument.OpenReader(new StringReader("Id,Padding,Value\n1,,ok\n2," + padding + ", \"bad\"\n"),
            options: new CsvReaderOptions { TrimWhitespace = true, QuoteParsingMode = CsvQuoteParsingMode.Strict });
        Assert.True(reader.Read());
        Assert.Equal("ok", reader.GetString(2));
        Assert.Throws<CsvParseException>(() => reader.Read());
    }

    [Theory]
    [InlineData(',', false, 0)]
    [InlineData(',', true, 0)]
    [InlineData('\u00fe', false, 0)]
    [InlineData('\u00fe', true, 0)]
    [InlineData('\u00ff', true, 0)]
    [InlineData(',', false, 1)]
    [InlineData(',', true, 1)]
    [InlineData('\u00fe', false, 1)]
    [InlineData('\u00fe', true, 1)]
    [InlineData('\u00ff', true, 1)]
    [InlineData(',', false, 2)]
    [InlineData(',', true, 2)]
    [InlineData('\u00fe', false, 2)]
    [InlineData('\u00fe', true, 2)]
    [InlineData('\u00ff', true, 2)]
    public void ReadersPreserveUnicodeAndQuotedFieldsAcrossPackedBlocks(char delimiter, bool quoted, int transport) {
        var text = new StringBuilder("Id" + delimiter + "Value\r\n");
        var expected = new string[257];
        for (int row = 0; row < expected.Length; row++) {
            // Cover every packed-block alignment, including Unicode whose low byte is punctuation.
            string value = new string('x', row % 32) + "\u0122\u010a\u800a\u800d\ufffe\u0100\uff00Żółć😀"
                + (char)(0x100 + delimiter) + " tail";
            if (quoted && row % 3 == 0) value += " \"quoted\"\r\nnext";
            if (row == 128) value += new string('z', 70_000);
            expected[row] = value;
            text.Append(row.ToString(CultureInfo.InvariantCulture)).Append(delimiter);
            if (quoted) text.Append('"').Append(value.Replace("\"", "\"\"")).Append('"');
            else text.Append(value);
            if (row < expected.Length - 1) text.Append(row % 2 == 0 ? "\r\n" : "\n");
        }

        var options = new CsvReaderOptions { Delimiter = delimiter, Culture = CultureInfo.InvariantCulture };
        using var bytes = new MemoryStream(Encoding.UTF8.GetBytes(text.ToString()));
        using TextReader input = transport == 0
            ? new StringReader(text.ToString())
            : new StreamReader(bytes, Encoding.UTF8);
        using var reader = transport == 1
            ? CsvDocument.OpenReader(bytes, options)
            : CsvDocument.OpenReader(input, options);
        string? retained = null;
        for (int row = 0; row < expected.Length; row++) {
            Assert.True(reader.Read());
            Assert.Equal(row, reader.GetInt32(0));
            Assert.Equal(expected[row], reader.GetString(1));
            if (row == 0) retained = reader.GetString(1);
        }
        Assert.False(reader.Read());
        Assert.Equal(expected[0], retained);
    }
}
