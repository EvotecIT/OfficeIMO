using System;
using System.IO;
using System.Linq;
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
}
