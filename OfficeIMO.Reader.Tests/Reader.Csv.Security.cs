using System.Text;
using OfficeIMO.Reader.Csv;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderCsvSecurityTests {
    [Fact]
    public void CsvReaderAdapter_RejectsRecordsBeyondConfiguredColumnBudgetBeforeNormalization() {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("a,b,c,d,e\n1,2,3,4,5\n"), writable: false);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            CsvReaderAdapter.Read(
                stream,
                sourceName: "wide.csv",
                csvOptions: new CsvReadOptions { MaxColumns = 4 }).ToList());

        Assert.Contains("5 columns", exception.Message, StringComparison.Ordinal);
    }
}
