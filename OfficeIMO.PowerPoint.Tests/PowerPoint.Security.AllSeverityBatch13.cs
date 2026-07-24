using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointAllSeverityBatch13SecurityTests {
    [Fact]
    public void AddTableStopsUnboundedSourceEnumerationAtConfiguredLimit() {
        using PowerPointPresentation presentation = PowerPointPresentation
            .Create();
        PowerPointSlide slide = presentation.AddSlide();

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => slide.AddTable(InfiniteRows(),
                options => options.MaxRows = 3));

        Assert.Contains("3-row", exception.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void AddTableStopsUnboundedNestedExpansionAtConfiguredLimit() {
        using PowerPointPresentation presentation = PowerPointPresentation
            .Create();
        PowerPointSlide slide = presentation.AddSlide();
        var rows = new[] {
            new NestedRow { Name = "item", Values = InfiniteValues() }
        };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => slide.AddTable(rows, options => {
                options.MaxRows = 3;
                options.CollectionMode = CollectionMode.ExpandRows;
            }));

        Assert.Contains("nested expansion", exception.Message,
            StringComparison.Ordinal);
    }

    private static IEnumerable<SimpleRow> InfiniteRows() {
        int value = 0;
        while (true) yield return new SimpleRow { Value = value++ };
    }

    private static IEnumerable<int> InfiniteValues() {
        int value = 0;
        while (true) yield return value++;
    }

    private sealed class SimpleRow {
        public int Value { get; set; }
    }

    private sealed class NestedRow {
        public string Name { get; set; } = string.Empty;
        public IEnumerable<int> Values { get; set; } = Array.Empty<int>();
    }
}
