using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class RecordConstructorMappingTests {
    [Fact]
    public void ReaderMapsPositionalRecordsThroughExistingColumnAliases() {
        using var data = new DataTable();
        data.Columns.Add("Order Id", typeof(int));
        data.Columns.Add("Customer name", typeof(string));
        data.Columns.Add("Total", typeof(string));
        data.Rows.Add(42, "Żółw", "12.50");
        using var reader = data.CreateDataReader();

        Assert.Equal(new Order(42, "Żółw", 12.50m), Assert.Single(reader.RowsAs<Order>()));
        Assert.False(reader.IsClosed);
    }

    [Fact]
    public void GetterOnlyRecordPropertiesUseTheirPrimaryConstructor() {
        using var data = new DataTable();
        data.Columns.Add("Name", typeof(string));
        data.Rows.Add(" north ");
        using var reader = data.CreateDataReader();
        Assert.Equal("NORTH", Assert.Single(reader.RowsAs<GetterOnlyRow>()).Name);
    }

    [Theory]
    [InlineData(7)]
    [InlineData(-1)]
    public void RecordStructConstructorsValidateAndNormalizePrimaryArguments(int value) {
        using var data = new DataTable();
        data.Columns.Add("Value", typeof(int));
        data.Rows.Add(value);
        using var reader = data.CreateDataReader();
        if (value < 0) {
            Assert.Throws<DataMappingException>(() => reader.RowsAs<ValidatedValue>().ToArray());
        } else {
            Assert.Equal(70, Assert.Single(reader.RowsAs<ValidatedValue>()).Value);
        }
    }

    [Theory]
    [InlineData(7, false)]
    [InlineData(-1, false)]
    [InlineData(7, true)]
    [InlineData(-1, true)]
    public void RecordStructRequiresMandatoryConstructorColumns(int value, bool includeUnit) {
        using var data = new DataTable();
        data.Columns.Add("Value", typeof(int));
        if (includeUnit) {
            data.Columns.Add("Unit", typeof(string));
            data.Rows.Add(value, "kg");
        } else {
            data.Rows.Add(value);
        }
        using var reader = data.CreateDataReader();
        if (!includeUnit || value < 0) {
            Assert.Throws<DataMappingException>(() => reader.RowsAs<MeasuredValue>().ToArray());
        } else {
            var row = Assert.Single(reader.RowsAs<MeasuredValue>());
            Assert.Equal(70, row.Value);
            Assert.Equal("kg", row.Unit);
        }
    }

    [Fact]
    public void PlainMutableStructStillUsesDefaultConstructionForPartialColumns() {
        using var data = new DataTable();
        data.Columns.Add("Value", typeof(int));
        data.Rows.Add(7);
        using var reader = data.CreateDataReader();
        var row = Assert.Single(reader.RowsAs<PlainValue>());
        Assert.Equal(7, row.Value);
        Assert.Null(row.Unit);
    }

    [Fact]
    public void ConstructorUsesCultureFormatsConverterAndDeclaredOptionalDefaults() {
        var plan = AutomaticRowMappingPlan<ConvertedOrder>.Create(["Amount EUR", "Day"]);
        object?[] values = ["EUR 12,50", "28/10/2026"];
        var row = plan.MapRow(index => values[index], CultureInfo.GetCultureInfo("fr-FR"), ["dd/MM/yyyy"],
            (value, type, culture) => type == typeof(decimal) && value is string text
                ? (true, decimal.Parse(text.Substring(4), culture)) : (false, null));

        Assert.Equal(12.50m, row.Amount);
        Assert.Equal(new DateTime(2026, 10, 28), row.Day);
        Assert.Equal("default", row.Label);
    }

    [Fact]
    public void ConstructorNormalizationIsNotOverwrittenByPropertyAssignment() {
        var plan = AutomaticRowMappingPlan<NormalizedOrder>.Create(["Customer"]);
        var row = plan.MapValues(["north"], CultureInfo.InvariantCulture, null);
        Assert.Equal("NORTH", row.Customer);
    }

    [Theory]
    [InlineData(DataMappingErrorValuePolicy.Redact)]
    [InlineData(DataMappingErrorValuePolicy.Include)]
    public void ConstructorValidationRespectsErrorValuePolicy(DataMappingErrorValuePolicy policy) {
        const string secret = "private order value";
        var plan = AutomaticRowMappingPlan<ValidatedOrder>.Create(["Value"]);
        DataMappingException error = Assert.Throws<DataMappingException>(() =>
            plan.MapValues([secret], CultureInfo.InvariantCulture, null, errorValuePolicy: policy));
        Assert.Contains(nameof(ValidatedOrder), error.Message);
        if (policy == DataMappingErrorValuePolicy.Redact) Assert.DoesNotContain(secret, error.ToString());
        else Assert.Contains(secret, error.Message);
        Assert.Null(error.InnerException);
    }

    [Fact]
    public void EquallyApplicableConstructorsRequireAnExplicitFactory() {
        Assert.Throws<DataMappingException>(() => AutomaticRowMappingPlan<AmbiguousOrder>.Create(["Value"]));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public async Task ReaderRecordsShareAsyncAndOrderedParallelMappingWithoutTakingOwnership() {
        using var data = new DataTable();
        data.Columns.Add("Value", typeof(int));
        data.Rows.Add(1);
        data.Rows.Add(2);
        using (var reader = data.CreateDataReader()) {
            var rows = new System.Collections.Generic.List<NumberRow>();
            await foreach (var row in reader.RowsAsAsync<NumberRow>()) rows.Add(row);
            Assert.Equal(new[] { new NumberRow(1), new NumberRow(2) }, rows);
            Assert.False(reader.IsClosed);
        }
        using (var reader = data.CreateDataReader()) {
            Assert.Equal(new[] { new NumberRow(1), new NumberRow(2) }, reader.RowsAsParallel<NumberRow>().ToArray());
            Assert.False(reader.IsClosed);
        }
        using var canceledReader = data.CreateDataReader();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () => {
            await foreach (var _ in canceledReader.RowsAsAsync<NumberRow>(cancellation.Token)) { }
        });
        Assert.False(canceledReader.IsClosed);
    }
#endif

    public sealed record Order([property: DataColumn("Order Id")] int Id,
        [property: DataColumn("Customer name")] string Customer, decimal Total);
    public sealed record ConvertedOrder([property: DataColumn("Amount EUR")] decimal Amount,
        DateTime Day, string Label = "default");
    public sealed record NumberRow(int Value);
    public sealed record GetterOnlyRow(string Name) {
        public string Name { get; } = Name.Trim().ToUpperInvariant();
    }
    public struct PlainValue {
        public int Value { get; set; }
        public string? Unit { get; set; }
    }
    public readonly record struct MeasuredValue(int Value, string Unit) {
        public int Value { get; init; } = Value < 0
            ? throw new ArgumentOutOfRangeException(nameof(Value)) : Value * 10;
    }
    public readonly record struct ValidatedValue(int Value) {
        public int Value { get; init; } = Value < 0
            ? throw new ArgumentOutOfRangeException(nameof(Value)) : Value * 10;
    }

    public sealed record NormalizedOrder {
        public string Customer { get; init; }
        public NormalizedOrder(string customer) => Customer = customer.ToUpperInvariant();
    }

    public sealed record ValidatedOrder {
        public string Value { get; init; }
        public ValidatedOrder(string value) => throw new ArgumentException(value);
    }

    public sealed class AmbiguousOrder {
        public int Value { get; init; }
        public AmbiguousOrder(int value, string label = "") => Value = value;
        public AmbiguousOrder(int value, int count = 0) => Value = value;
    }
}
