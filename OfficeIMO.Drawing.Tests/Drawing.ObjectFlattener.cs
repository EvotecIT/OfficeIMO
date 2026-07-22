using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingObjectFlattenerTests {
    [Fact]
    public void FlattenAppliesExplicitColumnsAndTheirOrder() {
        var options = new ObjectFlattenerOptions {
            Columns = new[] { "Email", "Name" }
        };

        Dictionary<string, object?> values = new ObjectFlattener().Flatten(new ContactRow(), options);

        Assert.Equal(new[] { "Email", "Name" }, values.Keys);
        Assert.DoesNotContain("Secret", values.Keys);
    }

    [Fact]
    public void FlattenAppliesIncludeAndExcludeFilters() {
        var options = new ObjectFlattenerOptions {
            IncludeProperties = new[] { "Name", "Email", "Secret" },
            ExcludeProperties = new[] { "Secret", "Email" }
        };

        Dictionary<string, object?> values = new ObjectFlattener().Flatten(new ContactRow(), options);

        Assert.Equal("Alice", Assert.Single(values).Value);
        Assert.True(values.ContainsKey("Name"));
    }

    [Fact]
    public void FlattenAndGetPathsRetainLongValueTupleElements() {
        var tuple = (1, 2, 3, 4, 5, 6, 7, 8, 9);
        var flattener = new ObjectFlattener();
        var options = new ObjectFlattenerOptions();

        Dictionary<string, object?> values = flattener.Flatten(tuple, options);
        List<string> paths = flattener.GetPaths(tuple.GetType(), options);

        Assert.Equal(9, values.Count);
        Assert.Equal(8, values["Item8"]);
        Assert.Equal(9, values["Item9"]);
        Assert.Equal(Enumerable.Range(1, 9).Select(index => $"Item{index}"), paths);
    }

    private sealed class ContactRow {
        public string Name { get; } = "Alice";

        public string Email { get; } = "alice@example.test";

        public string Secret { get; } = "private";
    }
}
