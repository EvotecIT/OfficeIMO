using System;
using System.Globalization;
using System.IO;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvSkippedRecordQuoteTests {
    [Theory]
    [InlineData(',', 1)]
    [InlineData(',', 4)]
    [InlineData(',', 5)]
    [InlineData('|', 1)]
    [InlineData('|', 4)]
    [InlineData('|', 5)]
    public void StrictReadersValidateQuotedFieldsInSkippedRecords(char delimiter, int quotedColumn) {
        string separator = delimiter.ToString();
        string[] fields = { new string('a', 40), "b", "c", "d", "e", "f" };
        fields[quotedColumn] = "\"skip, | \"\"quoted\"\"\"";
        string retainedRecord = "1" + separator + "2" + separator + "3" + separator + "4" + separator + "5" + separator + "6\n";
        var options = new CsvReaderOptions {
            Delimiter = delimiter,
            HasHeaderRow = false,
            SkipInitialRecords = 1,
            QuoteParsingMode = CsvQuoteParsingMode.Strict
        };
        using (var reader = CsvDocument.OpenReader(new StringReader(string.Join(separator, fields) + "\n" + retainedRecord), options)) {
            Assert.True(reader.Read());
            Assert.Equal(6, reader.FieldCount);
            for (int column = 0; column < 6; column++) {
                Assert.Equal((column + 1).ToString(CultureInfo.InvariantCulture), reader.GetString(column));
            }
            Assert.False(reader.Read());
        }

        fields[quotedColumn] = "\"bad\"x";
        Assert.Throws<CsvParseException>(() => {
            using var reader = CsvDocument.OpenReader(new StringReader(string.Join(separator, fields) + "\n" + retainedRecord), options);
            reader.Read();
        });
    }
}
