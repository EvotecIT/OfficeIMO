using OfficeIMO.CSV;
using System;
using System.ComponentModel;
using System.Globalization;
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
    public void ReaderRowsAs_InvalidAutomaticValueReportsTargetProperty() {
        using var reader = CsvDocument.OpenTextDataReader("Order Id\nnot-an-integer\n");

        DataMappingException exception = Assert.Throws<DataMappingException>(() =>
            reader.RowsAs<SalesRow>().ToArray());

        Assert.Contains("SalesRow.OrderId", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderRowsAs_DoesNotRetryAPropertySetterThatThrowsAConversionException() {
        ThrowingSetterRow.SetterCalls = 0;
        using var reader = CsvDocument.OpenTextDataReader("Value\n42\n");

        Assert.Throws<FormatException>(() => reader.RowsAs<ThrowingSetterRow>().ToArray());

        Assert.Equal(1, ThrowingSetterRow.SetterCalls);
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
    public void ReaderRowsAs_MapsSupportedTypedGettersWithoutChangingValues() {
        const string batchId = "2fae048c-5886-43ec-b03f-e5814c5d52ba";
        const string created = "2026-08-07T17:05:04.1234567Z";
        const string csv =
            "Text,Flag,ByteValue,Code,Created,Amount,DoubleValue,FloatValue,BatchId,ShortValue,IntValue,LongValue\n" +
            $"Alpha,true,255,Z,{created},12345.6789,1.25,2.5,{batchId},-1234,123456789,-9876543210\n";

        using var reader = CsvDocument.OpenTextDataReader(csv);
        TypedGetterRow row = Assert.Single(reader.RowsAs<TypedGetterRow>());

        Assert.Equal("Alpha", row.Text);
        Assert.True(row.Flag);
        Assert.Equal(byte.MaxValue, row.ByteValue);
        Assert.Equal('Z', row.Code);
        Assert.Equal(DateTimeKind.Utc, row.Created.Kind);
        Assert.Equal(DateTime.Parse(created, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind), row.Created);
        Assert.Equal(12345.6789m, row.Amount);
        Assert.Equal(1.25d, row.DoubleValue);
        Assert.Equal(2.5f, row.FloatValue);
        Assert.Equal(Guid.Parse(batchId), row.BatchId);
        Assert.Equal((short)-1234, row.ShortValue);
        Assert.Equal(123456789, row.IntValue);
        Assert.Equal(-9876543210L, row.LongValue);
    }

    [Theory]
    [InlineData("2026-08-07T17:05:04.1234567Z", DateTimeKind.Utc)]
    [InlineData("2026-08-07T17:05:04.1234567", DateTimeKind.Unspecified)]
    public void RowsAs_PreservesRoundTripDateTimeKind(string value, DateTimeKind expectedKind) {
        string csv = $"Created\n{value}\n";
        DateTime expected = DateTime.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);

        DateTimeRow materialized = Assert.Single(CsvDocument.Parse(csv).RowsAs<DateTimeRow>());
        using var reader = CsvDocument.OpenTextDataReader(csv);
        DateTimeRow forwardOnly = Assert.Single(reader.RowsAs<DateTimeRow>());

        Assert.Equal(expectedKind, materialized.Created.Kind);
        Assert.Equal(expectedKind, forwardOnly.Created.Kind);
        Assert.Equal(expected.Ticks, materialized.Created.Ticks);
        Assert.Equal(expected.Ticks, forwardOnly.Created.Ticks);
    }

    [Fact]
    public void RowsAs_PreservesRoundTripDateTimeWithExactFormat() {
        const string value = "2026-08-07T17:05:04.1234567Z";
        string csv = $"Created\n{value}\n";
        var options = new CsvLoadOptions { DateTimeFormats = new[] { "O" } };
        DateTime expected = DateTime.ParseExact(
            value,
            "O",
            CultureInfo.InvariantCulture,
            DateTimeStyles.RoundtripKind);

        DateTimeRow materialized = Assert.Single(CsvDocument.Parse(csv, options).RowsAs<DateTimeRow>());
        using var reader = CsvDocument.OpenTextDataReader(csv, options);
        DateTimeRow forwardOnly = Assert.Single(reader.RowsAs<DateTimeRow>());

        Assert.Equal(DateTimeKind.Utc, materialized.Created.Kind);
        Assert.Equal(DateTimeKind.Utc, forwardOnly.Created.Kind);
        Assert.Equal(expected.Ticks, materialized.Created.Ticks);
        Assert.Equal(expected.Ticks, forwardOnly.Created.Ticks);
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

    private sealed class ThrowingSetterRow {
        internal static int SetterCalls;

        public int Value {
            get => 0;
            set {
                SetterCalls++;
                throw new FormatException("Property setter failure.");
            }
        }
    }

    private sealed class DateTimeRow {
        public DateTime Created { get; set; }
    }

    private sealed class TypedGetterRow {
        public string Text { get; set; } = string.Empty;
        public bool Flag { get; set; }
        public byte ByteValue { get; set; }
        public char Code { get; set; }
        public DateTime Created { get; set; }
        public decimal Amount { get; set; }
        public double DoubleValue { get; set; }
        public float FloatValue { get; set; }
        public Guid BatchId { get; set; }
        public short ShortValue { get; set; }
        public int IntValue { get; set; }
        public long LongValue { get; set; }
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
