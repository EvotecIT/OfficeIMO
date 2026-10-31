using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeSeriesTests {
    [Fact]
    public void NumericSelectorsPreserveValuesAndInferPropertyNames() {
        var rows = new[] { new Values { Revenue = 12.5m, Count = 7, Ratio = 0.5f, Optional = 3 } };
        Assert.Equal("Revenue", OfficeSeries.From(rows, row => row.Revenue).Name);
        Assert.Equal(new[] { 12.5 }, OfficeSeries.From(rows, row => row.Revenue).Values);
        Assert.Equal(new[] { 7d }, OfficeSeries.From(rows, row => row.Count).Values);
        Assert.Equal(new[] { 0.5 }, OfficeSeries.From(rows, row => row.Ratio).Values);
        Assert.Equal(new[] { 3d }, OfficeSeries.From(rows, row => row.Optional).Values);
        Assert.Equal(new[] { 25d }, OfficeSeries.From(rows, row => row.Revenue * 2, "Projected").Values);
    }

    [Fact]
    public void InvalidOrMissingNumericValuesAreRejected() {
        var rows = new[] { new Values() };
        Assert.Throws<ArgumentException>(() => OfficeSeries.From(rows, row => row.Label));
        Assert.Throws<ArgumentException>(() => OfficeSeries.From(rows, row => row.Revenue + 1));
        Assert.Throws<ArgumentException>(() => OfficeSeries.From(rows, row => row.Optional));
    }

    private sealed class Values {
        public decimal Revenue { get; set; }
        public long Count { get; set; }
        public float Ratio { get; set; }
        public int? Optional { get; set; }
        public string Label { get; set; } = "label";
    }
}
