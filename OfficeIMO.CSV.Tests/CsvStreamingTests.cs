using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public class CsvStreamingTests
{
    [Fact]
    public void StreamingMode_ReadsLazily()
    {
        var tempPath = Path.GetTempFileName();
        try
        {
            File.WriteAllText(tempPath, "Id,Name\n1,Alice\n2,Bob\n3,Charlie\n");

            var options = new CsvLoadOptions { Mode = CsvLoadMode.Stream };
            var doc = CsvDocument.Load(tempPath, options);

            var rows = doc.AsEnumerable().ToList();
            Assert.Equal(3, rows.Count);
            Assert.Throws<InvalidOperationException>(() => doc.SortBy("Id"));

            doc.Materialize().SortBy("Id");
            Assert.Equal(1, doc.AsEnumerable().First().AsInt32("Id"));
        }
        finally
        {
            File.Delete(tempPath);
        }
    }

    [Fact]
    public void StreamingMode_Disallows_FilterUntilMaterialized()
    {
        var csv = "Id,Value\n1,A\n2,B\n";
        var doc = CsvDocument.Parse(csv, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        Assert.Throws<InvalidOperationException>(() => doc.Filter(_ => true));
    }

    [Fact]
    public void LoadFromStream_InMemoryMode_ParsesRows()
    {
        var bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n2,Bob\n");
        using var stream = new MemoryStream(bytes, writable: false);

        var doc = CsvDocument.Load(stream, new CsvLoadOptions { Mode = CsvLoadMode.InMemory }, leaveOpen: true);
        var rows = doc.AsEnumerable().ToList();

        Assert.Equal(2, rows.Count);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void LoadFromStream_StreamMode_CanReenumerateRows()
    {
        var bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n2,Bob\n");
        using var stream = new MemoryStream(bytes, writable: false);

        var doc = CsvDocument.Load(stream, new CsvLoadOptions { Mode = CsvLoadMode.Stream }, leaveOpen: true);
        var firstPass = doc.AsEnumerable().Select(r => r.AsString("Name")).ToArray();
        var secondPass = doc.AsEnumerable().Select(r => r.AsString("Name")).ToArray();

        Assert.Equal(new[] { "Alice", "Bob" }, firstPass);
        Assert.Equal(new[] { "Alice", "Bob" }, secondPass);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void LoadFromStream_StreamMode_WithLeaveOpenFalse_DisposesSource()
    {
        var bytes = Encoding.UTF8.GetBytes("Id,Name\n1,Alice\n");
        var stream = new MemoryStream(bytes, writable: false);

        var doc = CsvDocument.Load(stream, new CsvLoadOptions { Mode = CsvLoadMode.Stream }, leaveOpen: false);
        var rows = doc.AsEnumerable().ToList();

        Assert.Single(rows);
        Assert.Throws<ObjectDisposedException>(() => stream.ReadByte());
    }
}
