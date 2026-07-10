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
    public void FromItems_OrdersPublicPropertiesByMetadata() {
        var table = TabularDataTableBuilder.FromItems(new object?[] {
            new StablePropertyOrderRow { Second = "two", First = 1 }
        });

        Assert.Equal(new[] { "Second", "First" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
        Assert.Equal("two", table.Rows[0]["Second"]);
        Assert.Equal(1, table.Rows[0]["First"]);
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
            ProjectObject = (item, _) => {
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
    public void FromItems_UnwrapsRowsWhenHostProjectionReturnsNoColumns() {
        var unwrapped = new[] { new object(), new object() };
        var table = TabularDataTableBuilder.FromItems(new object?[] {
            new HostRow(1),
            new HostRow(2)
        }, new TabularDataOptions {
            UnwrapValue = item => item is HostRow row ? unwrapped[row.Number - 1] : item,
            ProjectObject = (item, _) => item is HostRow { Number: 1 }
                ? null
                : new Dictionary<string, object?>()
        });

        Assert.Equal(new[] { "Value" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
        Assert.Same(unwrapped[0], table.Rows[0]["Value"]);
        Assert.Same(unwrapped[1], table.Rows[1]["Value"]);
    }

    [Fact]
    public void FromItems_BypassesHostProjectionForNativeTablesAndEnumerableContainers() {
        var source = new DataTable("Source");
        source.Columns.Add("Name", typeof(string));
        source.Rows.Add("Alpha");
        var projectionCalls = 0;
        var options = new TabularDataOptions {
            CopyExistingDataTable = false,
            ProjectObject = (item, _) => {
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
    public void FromItems_FirstRowProjectionReceivesEstablishedColumns() {
        IReadOnlyList<string>? secondRowColumns = null;
        var table = TabularDataTableBuilder.FromItems(new object?[] {
            new HostRow(1),
            new HostRow(2)
        }, new TabularDataOptions {
            ColumnDiscoveryMode = TabularColumnDiscoveryMode.FirstRow,
            ProjectObject = (item, columns) => {
                var row = (HostRow)item!;
                if (row.Number == 2) {
                    secondRowColumns = columns;
                    if (columns == null || !columns.Contains("Value")) {
                        throw new InvalidOperationException("Established columns were not provided.");
                    }
                }

                return new Dictionary<string, object?> { ["Value"] = row.Number };
            }
        });

        Assert.NotNull(secondRowColumns);
        Assert.Equal(new[] { "Value" }, secondRowColumns);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(2, table.Rows[1]["Value"]);
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

    [Fact]
    public void FromItems_AllRowsPreservesNewColumnsBesideCaseVariantMatches() {
        var table = TabularDataTableBuilder.FromItems(new object?[] {
            new Dictionary<string, object?>(StringComparer.Ordinal) { ["Name"] = "Alpha" },
            new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                ["name"] = "Beta",
                ["Age"] = 42
            }
        }, new TabularDataOptions {
            ColumnDiscoveryMode = TabularColumnDiscoveryMode.AllRows
        });

        Assert.Equal(new[] { "Name", "Age" }, table.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToArray());
        Assert.Equal("Beta", table.Rows[1]["Name"]);
        Assert.Equal(42, table.Rows[1]["Age"]);
    }

    private sealed class HostRow {
        public HostRow(int number) => Number = number;

        public int Number { get; }
    }

    private sealed class StablePropertyOrderRow {
        public string Second { get; set; } = string.Empty;

        public int First { get; set; }
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
