using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvDialectEncodingTests
{
    /// <summary>
    /// A delimiter or quote that the output encoding cannot write is replaced by the same substitute as unencodable
    /// field text, so readers can no longer find the field boundaries. Writers reject such a dialect up front.
    /// </summary>
    [Theory]
    [InlineData('\uD800', '"', null, "utf-8")]
    [InlineData(',', '\uDC00', null, "utf-8")]
    [InlineData(',', '\'', '\uD800', "utf-8")]
    [InlineData('§', '"', null, "us-ascii")]
    public void Writers_RejectSeparatorsTheOutputEncodingCannotWrite(char delimiter, char quote, char? escape, string encodingName)
    {
        var options = new CsvSaveOptions
        {
            Delimiter = delimiter,
            QuoteCharacter = quote,
            EscapeCharacter = escape,
            Encoding = Encoding.GetEncoding(encodingName)
        };
        CsvDocument document = CsvDocument.Parse("A,B\nx?y,z\n");
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CsvDialect.{Guid.NewGuid():N}.csv");
        try
        {
            using var destination = new MemoryStream();
            Assert.Throws<ArgumentException>(() => document.Save(destination, options));
            Assert.Equal(0, destination.Length);
            Assert.Throws<ArgumentException>(() => document.Save(path, options));
            Assert.False(File.Exists(path));
            Assert.Throws<ArgumentException>(() => CsvRowWriter.CreateStream(new MemoryStream(), options));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void NonAsciiDelimiter_RoundTripsThroughUtf8()
    {
        var save = new CsvSaveOptions { Delimiter = '§', Encoding = new UTF8Encoding(false) };
        CsvDocument document = CsvDocument.Parse("A,B\n\"x§y\",�\n");

        using var output = new MemoryStream();
        document.Save(output, save);
        output.Position = 0;
        CsvDocument loaded = CsvDocument.Load(output, new CsvLoadOptions { Delimiter = '§', Encoding = new UTF8Encoding(false) });

        CsvRow row = Assert.Single(loaded.AsEnumerable());
        Assert.Equal("x§y", row.Get<string>("A"));
        Assert.Equal("�", row.Get<string>("B"));
    }

    [Fact]
    public void CallerTextWriter_IsValidatedAgainstItsOwnEncoding()
    {
        // Encoding applies to file and stream output; a supplied writer decides what it can hold.
        var options = new CsvSaveOptions { QuoteCharacter = '§', Encoding = Encoding.ASCII };
        using var text = new StringWriter();
        using (var writer = new CsvRowWriter(text, options, leaveOpen: true))
        {
            writer.WriteTextRow(new[] { "A" }, new[] { "x,y" });
        }

        Assert.Contains("§x,y§", text.ToString(), StringComparison.Ordinal);
    }
}
