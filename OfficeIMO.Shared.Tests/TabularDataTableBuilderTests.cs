using System.Collections.Generic;
using System.Data;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public class TabularDataTableBuilderTests {
    [Fact]
    public void FromItems_ConvertsObjectsToTypedColumns() {
        var table = TabularDataTableBuilder.FromItems(new object?[] {
            new { Id = 1, Name = "Alice" },
            new { Id = 2, Name = "Bob" }
        });

        Assert.Equal(typeof(int), table.Columns["Id"]!.DataType);
        Assert.Equal(typeof(string), table.Columns["Name"]!.DataType);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("Bob", table.Rows[1]["Name"]);
    }

    [Fact]
    public void FromItems_ExpandsSingleEnumerableInput() {
        var rows = new[] {
            new Dictionary<string, object?> { ["Id"] = 1 },
            new Dictionary<string, object?> { ["Id"] = 2 }
        };

        var table = TabularDataTableBuilder.FromItems(new object?[] { rows });

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[1]["Id"]);
    }

    [Fact]
    public void FromItems_ConvertsDataRows() {
        var source = new DataTable("Source");
        source.Columns.Add("Name", typeof(string));
        source.Rows.Add("Alpha");
        source.Rows.Add("Beta");

        var table = TabularDataTableBuilder.FromItems(new object?[] {
            source.Rows[0],
            source.Rows[1]
        }, new TabularDataOptions { TableName = "Copy" });

        Assert.Equal("Copy", table.TableName);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal("Beta", table.Rows[1]["Name"]);
    }

    [Fact]
    public void FromItems_UsesHostProjectionCallback() {
        var projectionCalls = 0;
        var table = TabularDataTableBuilder.FromItems(new object?[] { new HostRow(7) }, new TabularDataOptions {
            ProjectObject = item => {
                projectionCalls++;
                return new Dictionary<string, object?> {
                    ["Value"] = item is HostRow row ? row.Number : -1
                };
            }
        });

        Assert.Equal(1, projectionCalls);
        Assert.Equal(typeof(int), table.Columns["Value"]!.DataType);
        Assert.Equal(7, table.Rows[0]["Value"]);
    }

    [Fact]
    public void FromItems_BypassesHostProjectionForNativeTablesAndEnumerableContainers() {
        var source = new DataTable("Source");
        source.Columns.Add("Name", typeof(string));
        source.Rows.Add("Alpha");
        var projectionCalls = 0;
        var options = new TabularDataOptions {
            CopyExistingDataTable = false,
            ProjectObject = item => {
                projectionCalls++;
                return new Dictionary<string, object?> { ["Projected"] = item };
            }
        };

        var table = TabularDataTableBuilder.FromItems(new object?[] { source }, options);
        var empty = TabularDataTableBuilder.FromItems(new object?[] { Array.Empty<object?>() }, options);

        Assert.Same(source, table);
        Assert.Equal("Alpha", table.Rows[0]["Name"]);
        Assert.Empty(empty.Rows);
        Assert.Empty(empty.Columns);
        Assert.Equal(0, projectionCalls);
    }

    [Fact]
    public void FromItems_ConvertsReadOnlyDictionaryRows() {
        IReadOnlyDictionary<string, object?> row = new ReadOnlyRow(new Dictionary<string, object?> {
            ["Id"] = 5,
            ["Name"] = "Alice"
        });

        var table = TabularDataTableBuilder.FromItems(new object?[] { row });

        Assert.Equal(typeof(int), table.Columns["Id"]!.DataType);
        Assert.Equal("Alice", table.Rows[0]["Name"]);
    }

    [Fact]
    public void FromItems_PreservesExplicitNullScalarRowsWhenConfigured() {
        var table = TabularDataTableBuilder.FromItems(new object?[] { null, "Beta" }, new TabularDataOptions {
            PreserveNullRows = true
        });

        Assert.NotNull(table.Columns["Value"]);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(DBNull.Value, table.Rows[0]["Value"]);
        Assert.Equal("Beta", table.Rows[1]["Value"]);
    }

    [Fact]
    public void FromItems_PreservesNullReturnedByValueNormalizer() {
        var table = TabularDataTableBuilder.FromItems(new object?[] {
            new Dictionary<string, object?> { ["Secret"] = "remove-me" }
        }, new TabularDataOptions {
            NormalizeValue = _ => null
        });

        Assert.Equal(DBNull.Value, table.Rows[0]["Secret"]);
    }

    [Fact]
    public void FromItems_PreservesDictionaryLookupComparer() {
        var ordinalTable = TabularDataTableBuilder.FromItems(new object?[] {
            new Dictionary<string, object?>(StringComparer.Ordinal) { ["Name"] = "Alpha" },
            new Dictionary<string, object?>(StringComparer.Ordinal) { ["name"] = "Beta" }
        }, new TabularDataOptions {
            ColumnDiscoveryMode = TabularColumnDiscoveryMode.FirstRow
        });
        var ignoreCaseTable = TabularDataTableBuilder.FromItems(new object?[] {
            new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) { ["Name"] = "Alpha" },
            new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) { ["name"] = "Beta" }
        }, new TabularDataOptions {
            ColumnDiscoveryMode = TabularColumnDiscoveryMode.FirstRow
        });
        var distinctCaseTable = TabularDataTableBuilder.FromItems(new object?[] {
            new Dictionary<string, object?>(StringComparer.Ordinal) {
                ["Name"] = "Upper",
                ["name"] = "Lower"
            }
        }, new TabularDataOptions {
            ColumnDiscoveryMode = TabularColumnDiscoveryMode.FirstRow
        });

        Assert.Equal(DBNull.Value, ordinalTable.Rows[1]["Name"]);
        Assert.Equal("Beta", ignoreCaseTable.Rows[1]["Name"]);
        Assert.Equal(2, distinctCaseTable.Columns.Count);
        Assert.Equal("Upper", distinctCaseTable.Rows[0]["Name"]);
        Assert.Equal("Lower", distinctCaseTable.Rows[0]["name"]);
    }

    private sealed class HostRow {
        public HostRow(int number) => Number = number;

        public int Number { get; }
    }

    private sealed class ReadOnlyRow : IReadOnlyDictionary<string, object?> {
        private readonly Dictionary<string, object?> _values;

        internal ReadOnlyRow(Dictionary<string, object?> values) => _values = values;

        public IEnumerable<string> Keys => _values.Keys;

        public IEnumerable<object?> Values => _values.Values;

        public int Count => _values.Count;

        public object? this[string key] => _values[key];

        public bool ContainsKey(string key) => _values.ContainsKey(key);

        public bool TryGetValue(string key, out object? value) => _values.TryGetValue(key, out value);

        public IEnumerator<KeyValuePair<string, object?>> GetEnumerator() => _values.GetEnumerator();

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
