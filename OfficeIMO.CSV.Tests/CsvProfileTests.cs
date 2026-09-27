using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvProfileTests {
    [Theory]
    [InlineData("Id,Name\n1\n")]
    [InlineData("Id,Name\n1,Alice,extra\n")]
    [InlineData("Id,Name\n1,\"Alice\"tail\n")]
    [InlineData("Id,Id\n1,2\n")]
    public void StrictRejectsMalformedInputAcrossDocumentAndReader(string text) {
        Assert.ThrowsAny<CsvException>(() => CsvDocument.Parse(text, CsvProfiles.CreateLoadOptions(CsvProfile.Strict)));
        Assert.ThrowsAny<CsvException>(() => {
            using var reader = CsvDocument.OpenTextDataReader(text, CsvProfiles.CreateLoadOptions(CsvProfile.Strict));
            while (reader.Read()) reader.GetValues(new object[reader.FieldCount]);
        });
    }

    [Fact]
    public void CompatibilityRetainsForgivingRowsAndRenamesDuplicateHeaders() {
        var document = CsvDocument.Parse("Id,Id\n1\n2,3,ignored\n",
            CsvProfiles.CreateLoadOptions(CsvProfile.Compatibility));
        using var reader = document.CreateDataReader();
        Assert.Equal("Id_2", reader.GetName(1));
        Assert.True(reader.Read());
        Assert.Equal("1", reader.GetString(0));
        Assert.Equal("", reader.GetString(1));
        Assert.True(reader.Read());
        Assert.Equal("3", reader.GetString(1));
    }

    [Fact]
    public void StrictPreservesQuotedWhitespaceMultilineAndLiteralCommentRows() {
        var document = CsvDocument.Parse("Id,Value\r\n#record,\" space \r\nnext \"\r\n",
            CsvProfiles.CreateLoadOptions(CsvProfile.Strict));
        using var reader = document.CreateDataReader();
        Assert.True(reader.Read());
        Assert.Equal("#record", reader.GetString(0));
        Assert.Equal(" space \r\nnext ", reader.GetString(1));
        Assert.False(reader.Read());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StrictRejectsInvalidUtf8EvenWithMatchingBom(bool bom) {
        byte[] bytes = (bom ? new byte[] { 0xEF, 0xBB, 0xBF } : Array.Empty<byte>())
            .Concat(Encoding.UTF8.GetBytes("Id,Unused\n1,"))
            .Concat(new byte[] { 0xC3, 0x28, 0x0A }).ToArray();
        using var stream = new MemoryStream(bytes);
        Assert.Throws<DecoderFallbackException>(() => CsvDocument.Load(stream, CsvProfiles.CreateLoadOptions(CsvProfile.Strict)));
        string path = Path.GetTempFileName();
        try {
            File.WriteAllBytes(path, bytes);
            Assert.Throws<DecoderFallbackException>(() => {
                using var reader = CsvDocument.OpenDataReader(path, CsvProfiles.CreateLoadOptions(CsvProfile.Strict));
                while (reader.Read()) reader.GetInt32(0);
            });
        } finally { File.Delete(path); }
    }

    [Fact]
    public void StrictAcceptsUtf8BomAndSpreadsheetDetectsUtf16AndDelimiter() {
        byte[] utf8 = new UTF8Encoding(true).GetPreamble().Concat(Encoding.UTF8.GetBytes("Id,Name\r\n1,Żółć\r\n")).ToArray();
        using var utf8Stream = new MemoryStream(utf8);
        using var strictReader = CsvDocument.Load(utf8Stream, CsvProfiles.CreateLoadOptions(CsvProfile.Strict)).CreateDataReader();
        Assert.True(strictReader.Read());
        Assert.Equal("Żółć", strictReader.GetString(1));
        byte[] utf16 = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes("Id;Name\r\n2;Żółć\r\n")).ToArray();
        using var utf16Stream = new MemoryStream(utf16);
        var spreadsheet = CsvDocument.Load(utf16Stream, CsvProfiles.CreateLoadOptions(CsvProfile.Spreadsheet));
        Assert.Equal(";", spreadsheet.DelimiterText);
        using var spreadsheetReader = spreadsheet.CreateDataReader();
        Assert.True(spreadsheetReader.Read());
        Assert.Equal("Żółć", spreadsheetReader.GetString(1));
        using var wrongEncoding = new MemoryStream(utf16);
        Assert.Throws<DecoderFallbackException>(() => CsvDocument.Load(wrongEncoding, CsvProfiles.CreateLoadOptions(CsvProfile.Strict)));
    }

    [Fact]
    public void StrictFastReaderFallbackPreservesLiteralBomAtLaterRecordBoundary() {
        string path = Path.GetTempFileName();
        try {
            File.WriteAllBytes(path, Encoding.UTF8.GetBytes("Id,Value\n1,plain\n\uFEFF2,\"quoted\"\n"));
            using var reader = CsvDocument.OpenDataReader(path, CsvProfiles.CreateLoadOptions(CsvProfile.Strict));
            Assert.True(reader.Read());
            Assert.True(reader.Read());
            Assert.Equal("\uFEFF2", reader.GetString(0));
            Assert.Equal("quoted", reader.GetString(1));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FixedNoPreambleEncodingHasSameHeaderOnDocumentAndFastReaderPaths(bool useDefaultEncoding) {
        byte[] bytes = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes("Id,Value\n1,plain\n")).ToArray();
        var options = new CsvLoadOptions { Encoding = useDefaultEncoding ? null : new UTF8Encoding(false), DetectEncodingFromByteOrderMarks = false };
        using var source = new MemoryStream(bytes);
        var document = CsvDocument.Load(source, options);
        Assert.Equal("\uFEFFId", document.Header[0]);
        string path = Path.GetTempFileName();
        try {
            File.WriteAllBytes(path, bytes);
            using var reader = CsvDocument.OpenDataReader(path, options);
            Assert.Equal(document.Header[0], reader.GetName(0));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(CsvProfile.Strict, false)]
    [InlineData(CsvProfile.Spreadsheet, true)]
    public void SaveProfilesKeepTypedNumbersAndApplyTextFormulaPolicy(CsvProfile profile, bool escape) {
        var document = new CsvDocument().WithHeader("Formula", "Number", "Text", "Unicode");
        document.AddRow("=1+2", -12, "  @SUM(A1)", "Żółć");
        byte[] bytes = document.ToBytes(CsvProfiles.CreateSaveOptions(profile));
        Assert.Equal(escape, bytes.Take(3).SequenceEqual(new byte[] { 0xEF, 0xBB, 0xBF }));
        string text = Encoding.UTF8.GetString(bytes).TrimStart('\uFEFF');
        Assert.Contains("\r\n", text);
        using var reader = CsvDocument.Parse(text).CreateDataReader();
        Assert.True(reader.Read());
        Assert.Equal(escape ? "'=1+2" : "=1+2", reader.GetString(0));
        Assert.Equal("-12", reader.GetString(1));
        Assert.Equal(escape ? "'  @SUM(A1)" : "  @SUM(A1)", reader.GetString(2));
        Assert.Equal("Żółć", reader.GetString(3));
    }

    [Fact]
    public void CreatedOptionsAreIndependentAndRemainCustomizable() {
        var first = CsvProfiles.CreateLoadOptions(CsvProfile.Strict);
        first.Delimiter = ';';
        first.QuoteParsingMode = CsvQuoteParsingMode.Lenient;
        var second = CsvProfiles.CreateLoadOptions(CsvProfile.Strict);
        Assert.ThrowsAny<CsvException>(() => CsvDocument.Parse("A,B\n1\n", second));
        using var reader = CsvDocument.Parse("A;B\n1;2\n", first).CreateDataReader();
        Assert.True(reader.Read());
        Assert.Equal("2", reader.GetString(1));
        Assert.Throws<ArgumentOutOfRangeException>(() => CsvProfiles.CreateLoadOptions((CsvProfile)99));
        Assert.Throws<ArgumentOutOfRangeException>(() => CsvProfiles.CreateSaveOptions((CsvProfile)99));
    }

#if NET8_0_OR_GREATER
    [Theory]
    [InlineData("A,B\n1\n")]
    [InlineData("A,B\n1,\"bad\"tail\n")]
    public async Task IncrementalStrictReaderRejectsMalformedRows(string text) {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(text));
        await Assert.ThrowsAnyAsync<CsvException>(async () => {
            using var reader = await CsvDocument.OpenStreamingDataReaderAsync(source, CsvProfiles.CreateLoadOptions(CsvProfile.Strict));
            while (await reader.ReadAsync()) { }
        });
        Assert.True(source.CanRead);
    }
#endif
}
