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

    [Fact]
    public void RowsAs_BindsTheCurrentSchemaWhenEnumerationBegins() {
        CsvDocument document = CsvDocument.Parse(
            "Order Id,Sales Channel,Amount\n42,Online,165258.24\n");
        var rows = document.RowsAs<SalesRow>();

        document.RemoveColumn("Sales Channel");

        SalesRow row = Assert.Single(rows);
        Assert.Equal(42, row.OrderId);
        Assert.Equal(string.Empty, row.SalesChannel);
        Assert.Equal(165258.24m, row.Amount);
    }

    [Fact]
    public void RowsAs_MapsOnlyPropertiesWithPublicSetters() {
        CsvDocument document = CsvDocument.Parse("Writable,InitOnly,PrivateSetter\n42,84,126\n");

        MixedAccessRow row = Assert.Single(document.RowsAs<MixedAccessRow>());

        Assert.Equal(42, row.Writable);
        Assert.Equal(84, row.InitOnly);
        Assert.Equal(7, row.PrivateSetter);
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

    private sealed class MixedAccessRow {
        public int Writable { get; set; }
        public int InitOnly { get; init; }
        public int PrivateSetter { get; private set; } = 7;
    }
}
