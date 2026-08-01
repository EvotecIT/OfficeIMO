using OfficeIMO.CSV;
using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvTypedRowsApiTests {
    [Fact]
    public void RowsAs_MapsHeadersWithoutRequiringAConfigurationDelegate() {
        CsvDocument document = CsvDocument.Parse(
            "Order Id,Sales Channel,Amount\n42,Online,165258.24\n");

        SalesRow row = Assert.Single(document.RowsAs<SalesRow>());

        Assert.Equal(42, row.OrderId);
        Assert.Equal("Online", row.SalesChannel);
        Assert.Equal(165258.24m, row.Amount);
    }

    [Fact]
    public void RowsAs_MapsWritableStructProperties() {
        CsvDocument document = CsvDocument.Parse("Order Id,Amount\n42,165258.24\n");

        SalesValue row = Assert.Single(document.RowsAs<SalesValue>());

        Assert.Equal(42, row.OrderId);
        Assert.Equal(165258.24m, row.Amount);
    }

    private sealed class SalesRow {
        public int OrderId { get; set; }
        public string SalesChannel { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }

    private struct SalesValue {
        public int OrderId { get; set; }
        public decimal Amount { get; set; }
    }
}
