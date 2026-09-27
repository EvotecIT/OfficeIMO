using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvRoundTripAndAppendTests
{
    [Theory]
    [InlineData(false, "")]
    [InlineData(true, "")]
    [InlineData(false, "\n")]
    [InlineData(true, "\n")]
    [InlineData(false, "\r")]
    [InlineData(true, "\r")]
    [InlineData(false, "\r\n")]
    [InlineData(true, "\r\n")]
    public async Task Append_Preserves_Record_Boundaries(bool asynchronous, string ending)
    {
        string path = TemporaryPath();
        try
        {
            File.WriteAllText(path, "A,B\n1,2" + ending, new UTF8Encoding(false));
            var document = CsvDocument.Parse("A,B\n3,4\n");
            var options = new CsvSaveOptions { Append = true, IncludeHeader = false, NewLine = "\n" };
            if (asynchronous) await document.SaveAsync(path, options);
            else document.Save(path, options);
            Assert.Equal("A,B\n1,2" + (ending.Length == 0 ? "\n" : ending) + "3,4\n", File.ReadAllText(path));
            Assert.Equal(2, CsvDocument.Load(path).AsEnumerable().Count());
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false, "utf8")]
    [InlineData(true, "utf8")]
    [InlineData(false, "utf16")]
    [InlineData(true, "utf16")]
    [InlineData(false, "utf16be")]
    [InlineData(true, "utf16be")]
    [InlineData(false, "utf32")]
    [InlineData(true, "utf32")]
    public async Task Append_Uses_Encoded_Boundary_Without_An_Interior_Bom(bool asynchronous, string encodingName)
    {
        Encoding encoding = encodingName switch {
            "utf16" => Encoding.Unicode,
            "utf16be" => Encoding.BigEndianUnicode,
            "utf32" => Encoding.UTF32,
            _ => new UTF8Encoding(true)
        };
        string path = TemporaryPath();
        try
        {
            File.WriteAllText(path, "A,B\n東京,2", encoding);
            var options = new CsvSaveOptions { Append = true, IncludeHeader = false, NewLine = "\n", Encoding = encoding };
            var document = new CsvDocument().WithHeader("A", "B").AddRow("Łódź", 4);
            if (asynchronous) await document.SaveAsync(path, options);
            else document.Save(path, options);
            string expected = "A,B\n東京,2\nŁódź,4\n";
            Assert.Equal(expected, File.ReadAllText(path, encoding));
            Assert.Equal(encoding.GetPreamble().Concat(encoding.GetBytes(expected)).ToArray(), File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("objects")]
    [InlineData("reader")]
    [InlineData("parallel")]
    public void Export_Paths_Share_Append_Boundary_Handling(string kind)
    {
        string path = TemporaryPath();
        try
        {
            File.WriteAllText(path, "A,B\n1,2");
            var options = new CsvSaveOptions { Append = true, IncludeHeader = false, NewLine = "\n" };
            if (kind == "objects")
                CsvDocument.SaveObjects(path, new object[] { new { A = 3, B = 4 } }, options);
            else
            {
                using var table = new DataTable();
                table.Columns.Add("A", typeof(int));
                table.Columns.Add("B", typeof(int));
                table.Rows.Add(3, 4);
                using var reader = table.CreateDataReader();
                if (kind == "parallel") CsvDocument.WriteDataReaderParallel(path, reader, options);
                else CsvDocument.WriteDataReader(path, reader, options);
            }
            Assert.Equal("A,B\n1,2\n3,4\n", File.ReadAllText(path));
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Empty_Append_Does_Not_Modify_An_Unterminated_Record(bool asynchronous)
    {
        string path = TemporaryPath();
        try
        {
            File.WriteAllText(path, "A,B\n1,2");
            byte[] expected = File.ReadAllBytes(path);
            var document = new CsvDocument().WithHeader("A", "B");
            var options = new CsvSaveOptions { Append = true, IncludeHeader = false };
            if (asynchronous) await document.SaveAsync(path, options);
            else document.Save(path, options);
            Assert.Equal(expected, File.ReadAllBytes(path));
        }
        finally { File.Delete(path); }
    }

    [Theory]
    [InlineData("\t")]
    [InlineData(";")]
    [InlineData("||")]
    [InlineData("")]
    public void Default_Save_And_Readers_Retain_The_Complete_Delimiter(string delimiter)
    {
        string effectiveDelimiter = delimiter.Length == 0 ? "," : delimiter;
        string text = "A" + effectiveDelimiter + "B\r\n東京" + effectiveDelimiter + "4\r\n";
        CsvDocument document = CsvDocument.Parse(text, new CsvLoadOptions { DelimiterText = delimiter });
        Assert.Equal(effectiveDelimiter, document.DelimiterText);
        Assert.Equal(effectiveDelimiter[0], document.Delimiter);
        Assert.Equal(text, document.ToString());
        Assert.Equal(text, Encoding.UTF8.GetString(document.ToBytes()));
        using var reader = document.CreateDataReader();
        Assert.Equal(effectiveDelimiter, Assert.IsAssignableFrom<ICsvDataReaderDialectMetadata>(reader).DelimiterText);
        using var textReader = CsvDocument.OpenTextDataReader(text, new CsvLoadOptions { DelimiterText = delimiter });
        Assert.Equal(effectiveDelimiter, Assert.IsAssignableFrom<ICsvDataReaderDialectMetadata>(textReader).DelimiterText);
        Assert.True(reader.Read());
        Assert.Equal("4", reader.GetString(1));
        Assert.Equal("A,B\r\n東京,4\r\n", document.WithDelimiter(',').ToString());
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Append." + Guid.NewGuid().ToString("N") + ".csv");
}
