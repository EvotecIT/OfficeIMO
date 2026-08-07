using OfficeIMO.CSV;
using System;
using System.ComponentModel;
using System.Linq;
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

    [Fact]
    public void RowsAs_ReservesExactHeadersBeforeFriendlyFallbacks() {
        CsvDocument document = CsvDocument.Parse("Order Id,Order_Id\n42,84\n");

        ExactPriorityRow row = Assert.Single(document.RowsAs<ExactPriorityRow>());

        Assert.Equal(42, row.OrderId);
        Assert.Equal(84, row.Order_Id);
    }

    [Fact]
    public void RowsAs_RejectsAmbiguousFriendlyPropertyMatches() {
        CsvDocument document = CsvDocument.Parse("Order Id\n42\n");

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            document.RowsAs<ExactPriorityRow>().ToArray());

        Assert.Contains("matches multiple writable properties", exception.Message);
    }

    [Fact]
    public void RowsAs_IgnoresUnusedHiddenPropertyAmbiguity() {
        CsvDocument document = CsvDocument.Parse("Other\n42\n");

        HiddenDerivedRow row = Assert.Single(document.RowsAs<HiddenDerivedRow>());

        Assert.Equal(42, row.Other);
    }

    [Fact]
    public void RowsAs_RejectsUsedHiddenPropertyAmbiguity() {
        CsvDocument document = CsvDocument.Parse("Hidden\n42\n");

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            document.RowsAs<HiddenDerivedRow>().ToArray());

        Assert.Contains("with that exact name", exception.Message);
    }

    [Fact]
    public void ReaderRowsAs_MapsForwardOnlyRowsWithoutMaterializingADocument() {
        using var reader = CsvDocument.OpenTextDataReader(
            "Order Id,Sales Channel,Amount\n42,Online,165258.24\n84,Partner,12.50\n");

        SalesRow[] rows = reader.RowsAs<SalesRow>().ToArray();

        Assert.Equal(2, rows.Length);
        Assert.Equal(42, rows[0].OrderId);
        Assert.Equal("Partner", rows[1].SalesChannel);
        Assert.Equal(12.50m, rows[1].Amount);
        Assert.False(reader.IsClosed);
    }

    [Fact]
    public void ReaderRowsAs_UsesCsvReaderCulture() {
        var options = new CsvLoadOptions {
            Delimiter = ';',
            Culture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL")
        };
        using var reader = CsvDocument.OpenTextDataReader("Order Id;Amount\n42;165258,24\n", options);

        SalesRow row = Assert.Single(reader.RowsAs<SalesRow>());

        Assert.Equal(165258.24m, row.Amount);
    }

    [Fact]
    public void ReaderRowsAs_ReusesExplicitAotFriendlyDocumentMapping() {
        const string csv = "Order Id,Amount\n42,165258.24\n";
        Action<RowMapper<SalesValue>> mapping = map => map
            .FromColumn<int>("Order Id", static (item, value) => { item.OrderId = value; return item; })
            .FromColumn<decimal>("Amount", static (item, value) => { item.Amount = value; return item; });

        SalesValue materialized = Assert.Single(CsvDocument.Parse(csv).RowsAs(mapping));
        using var reader = CsvDocument.OpenTextDataReader(csv);
        SalesValue forwardOnly = Assert.Single(reader.RowsAs(mapping));

        Assert.Equal(materialized.OrderId, forwardOnly.OrderId);
        Assert.Equal(materialized.Amount, forwardOnly.Amount);
    }

    [Fact]
    public void RowsAs_UsesStandardDeclaredColumnAliases() {
        CsvDocument document = CsvDocument.Parse("Order Number,Amount\n42,165258.24\n");

        AliasedSalesRow row = Assert.Single(document.RowsAs<AliasedSalesRow>());

        Assert.Equal(42, row.OrderId);
        Assert.Equal(165258.24m, row.Amount);
    }

#if NET6_0_OR_GREATER
    [Fact]
    public void RowsAs_ConvertsExplicitDateOnlyAndTimeOnlyTargetsWithoutChangingInference() {
        CsvDocument document = CsvDocument.Parse("Date,Time\n2026-08-06,14:35:12\n");

        DateAndTimeRow row = Assert.Single(document.RowsAs<DateAndTimeRow>());
        CsvSchema inferred = document.InferSchema();

        Assert.Equal(new DateOnly(2026, 8, 6), row.Date);
        Assert.Equal(new TimeOnly(14, 35, 12), row.Time);
        Assert.Equal(typeof(DateTime), inferred.Columns[0].DataType);
    }
#endif

    [Fact]
    public void RowsAs_RedactsSourceValuesWhenRequested() {
        const string secret = "customer-secret-value";
        CsvDocument document = CsvDocument.Parse(
            $"Order Id\n{secret}\n",
            new CsvLoadOptions { MappingErrorValuePolicy = DataMappingErrorValuePolicy.Redact });

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            document.RowsAs<SalesRow>().ToArray());

        Assert.DoesNotContain(secret, exception.ToString(), StringComparison.Ordinal);
        Assert.Contains("cannot be converted", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RowsAs_PreservesSourceValuesByDefault() {
        const string sourceValue = "not-an-order-id";
        CsvDocument document = CsvDocument.Parse($"Order Id\n{sourceValue}\n");

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            document.RowsAs<SalesRow>().ToArray());

        Assert.Contains(sourceValue, exception.ToString(), StringComparison.Ordinal);
    }

#if NET6_0_OR_GREATER
    private sealed class DateAndTimeRow {
        public DateOnly Date { get; set; }
        public TimeOnly Time { get; set; }
    }
#endif

    private sealed class SalesRow {
        public int OrderId { get; set; }
        public string SalesChannel { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }

    private sealed class AliasedSalesRow {
        [DisplayName("Order Number")]
        public int OrderId { get; set; }
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

    private sealed class ExactPriorityRow {
        public int OrderId { get; set; }
        public int Order_Id { get; set; }
    }

    private class HiddenBaseRow {
        public int Hidden { get; set; }
    }

    private sealed class HiddenDerivedRow : HiddenBaseRow {
        public new string Hidden { get; set; } = string.Empty;
        public int Other { get; set; }
    }
}
