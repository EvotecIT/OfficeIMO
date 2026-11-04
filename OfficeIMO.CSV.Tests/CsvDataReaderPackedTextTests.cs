using System;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDataReaderPackedTextTests {
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
    [InlineData(',', false)]
    [InlineData(',', true)]
    [InlineData('\u00fe', false)]
    [InlineData('\u00fe', true)]
    [InlineData('\u00ff', true)]
    public void SnapshotTextPreservesUnicodeAndQuotedFieldsAcrossPackedBlocks(char delimiter, bool quoted) {
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

        using var reader = CsvDocument.OpenReader(new StringReader(text.ToString()),
            options: new CsvReaderOptions { Delimiter = delimiter, Culture = CultureInfo.InvariantCulture });
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
