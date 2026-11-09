using System.Data;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class DataReaderMappingFallbackTests {
    [Fact]
    public void ExplicitParallelMappingConfiguresOnceForObjectSchemaFallback() {
        using var table = new DataTable();
        table.Columns.Add("Payload", typeof(object));
        var first = new List<int> { 1, 2 };
        var second = new List<int> { 3, 4 };
        table.Rows.Add(new object[] { first });
        table.Rows.Add(new object[] { second });
        using var reader = table.CreateDataReader();
        int configureCalls = 0;

        PayloadRow[] rows = reader.RowsAsParallel<PayloadRow>(map => {
            configureCalls++;
            map.FromColumn<object>("Payload", static (row, value) => {
                row.Payload = value;
                return row;
            });
        }, new ParallelRowMappingOptions { MaxDegreeOfParallelism = 2, BatchSize = 1 }).ToArray();

        Assert.Collection(rows,
            row => Assert.Same(first, row.Payload),
            row => Assert.Same(second, row.Payload));
        Assert.Equal(1, configureCalls);
    }

    private sealed class PayloadRow {
        public PayloadRow() { }
        public object? Payload { get; set; }
    }
}
