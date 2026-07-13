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

    private sealed class ContactRow {
        public string Name { get; } = "Alice";

        public string Email { get; } = "alice@example.test";

        public string Secret { get; } = "private";
    }
}
