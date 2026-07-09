using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects() {
            var items = new[] {
                new { Name = "A", Value = 1 },
                new { Name = "B", Value = 2 }
            };

            var table = ObjectDataTableBuilder.FromObjects(items, "Data");

            Assert.Equal(2, table.Columns.Count);
            Assert.Equal("Name", table.Columns[0].ColumnName);
            Assert.Equal("Value", table.Columns[1].ColumnName);
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal("A", table.Rows[0]["Name"]);
            Assert.Equal(1, table.Rows[0]["Value"]);
            Assert.Equal("B", table.Rows[1]["Name"]);
            Assert.Equal(2, table.Rows[1]["Value"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromReadOnlyList_DoesNotSnapshotEnumerate() {
            var items = new ThrowOnEnumerateReadOnlyList<object?>(
                new ObjectDataBaseRow { Name = "A", Value = 1 },
                new ObjectDataBaseRow { Name = "B", Value = 2 });

            var table = ObjectDataTableBuilder.FromObjects(items, "Data");

            Assert.Equal(2, table.Rows.Count);
            Assert.Equal("A", table.Rows[0]["Name"]);
            Assert.Equal(1, table.Rows[0]["Value"]);
            Assert.Equal("B", table.Rows[1]["Name"]);
            Assert.Equal(2, table.Rows[1]["Value"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromDictionaries_PreservesColumnsAndDbNulls() {
            var items = new[] {
                new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                    ["Name"] = "A",
                    ["Value"] = 1,
                    ["Notes"] = null
                },
                new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                    ["name"] = "B",
                    ["Value"] = 2
                }
            };

            var table = ObjectDataTableBuilder.FromObjects(items, "Data");

            Assert.Equal(new[] { "Name", "Value", "Notes" }, table.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName).ToArray());
            Assert.Equal("A", table.Rows[0]["Name"]);
            Assert.Equal(1, table.Rows[0]["Value"]);
            Assert.Equal(DBNull.Value, table.Rows[0]["Notes"]);
            Assert.Equal("B", table.Rows[1]["Name"]);
            Assert.Equal(2, table.Rows[1]["Value"]);
            Assert.Equal(DBNull.Value, table.Rows[1]["Notes"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_JoinsCollectionValuesForCellDisplay() {
            var items = new object?[] {
                new Dictionary<string, object?> {
                    ["Name"] = "A",
                    ["Tags"] = new[] { "one", "two" }
                },
                new Dictionary<string, object?> {
                    ["Name"] = "B",
                    ["Tags"] = new List<int> { 1, 2, 3 }
                }
            };

            var table = ObjectDataTableBuilder.FromObjects(items, "Data");

            Assert.Equal("one, two", table.Rows[0]["Tags"]);
            Assert.Equal("1, 2, 3", table.Rows[1]["Tags"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromSparseOrdinalDictionaries_PreservesBlanks() {
            var first = new Dictionary<string, object?>();
            for (int i = 0; i < 32; i++) {
                first["Column" + i] = i == 0 ? "First" : null;
            }

            var second = new Dictionary<string, object?> {
                ["Column0"] = "Second",
                ["Column31"] = 31
            };

            var table = ObjectDataTableBuilder.FromObjects(new object?[] { first, second }, "Data");

            Assert.Equal(32, table.Columns.Count);
            Assert.Equal("First", table.Rows[0]["Column0"]);
            Assert.Equal(DBNull.Value, table.Rows[0]["Column31"]);
            Assert.Equal("Second", table.Rows[1]["Column0"]);
            Assert.Equal(DBNull.Value, table.Rows[1]["Column1"]);
            Assert.Equal(31, table.Rows[1]["Column31"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromCaseInsensitiveDictionaries_UsesDictionaryComparer() {
            var items = new[] {
                new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                    ["Name"] = "A",
                    ["Value"] = 1
                },
                new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) {
                    ["name"] = "B",
                    ["value"] = 2
                }
            };

            var table = ObjectDataTableBuilder.FromObjects(items, "Data");

            Assert.Equal("B", table.Rows[1]["Name"]);
            Assert.Equal(2, table.Rows[1]["Value"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromSparseHashtables_PreservesCaseInsensitiveLookupAndExactPrecedence() {
            var first = new Hashtable();
            for (int i = 0; i < 40; i++) {
                first["Column" + i] = i;
            }

            var second = new Hashtable {
                ["column0"] = "Insensitive",
                ["Column0"] = "Exact",
                ["column39"] = 39
            };

            var table = ObjectDataTableBuilder.FromObjects(new object?[] { first, second }, "Data");

            Assert.Equal(40, table.Columns.Count);
            Assert.Equal("Exact", table.Rows[1]["Column0"]);
            Assert.Equal(DBNull.Value, table.Rows[1]["Column1"]);
            Assert.Equal(39, table.Rows[1]["Column39"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects_InheritedProperties() {
            var items = new InheritedObjectRow[] {
                new() { Name = "A", Value = 1, Notes = "First" },
                new() { Name = "B", Value = 2, Notes = "Second" }
            };

            var table = ObjectDataTableBuilder.FromObjects(items, "Data");

            Assert.Contains("Name", table.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName));
            Assert.Contains("Value", table.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName));
            Assert.Contains("Notes", table.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName));
            Assert.Equal("A", table.Rows[0]["Name"]);
            Assert.Equal(1, table.Rows[0]["Value"]);
            Assert.Equal("First", table.Rows[0]["Notes"]);
            Assert.Equal("B", table.Rows[1]["Name"]);
            Assert.Equal(2, table.Rows[1]["Value"]);
            Assert.Equal("Second", table.Rows[1]["Notes"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects_HiddenDerivedPropertiesUseRuntimeMember() {
            ObjectDataTableBuilder.FromObjects(new ObjectDataBaseRow[] {
                new() { Name = "Warm", Value = 0 }
            }, "Warmup");

            var items = new ObjectDataBaseRow[] {
                new() { Name = "Base", Value = 1 },
                new HiddenObjectRow { Name = "Derived", Value = 2 }
            };

            var table = ObjectDataTableBuilder.FromObjects(items, "Data");

            Assert.Equal("Base", table.Rows[0]["Name"]);
            Assert.Equal(1, table.Rows[0]["Value"]);
            Assert.Equal("Derived", table.Rows[1]["Name"]);
            Assert.Equal(2, table.Rows[1]["Value"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromSingleEnumerableObject_PreservesTheRow() {
            var row = new EnumerableObjectRow { Name = "Single", Value = 42 };

            var table = ObjectDataTableBuilder.FromObjects(new object?[] { row }, "Data");

            Assert.Single(table.Rows.Cast<System.Data.DataRow>());
            Assert.Equal("Single", table.Rows[0]["Name"]);
            Assert.Equal(42, table.Rows[0]["Value"]);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects_ThrowsOnNullItems() {
            var ex = Assert.Throws<ArgumentNullException>(() => ObjectDataTableBuilder.FromObjects(null!));
            Assert.Equal("items", ex.ParamName);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects_ThrowsOnEmptyItems() {
            var ex = Assert.Throws<ArgumentException>(() => ObjectDataTableBuilder.FromObjects(Array.Empty<object?>()));
            Assert.StartsWith("Provide at least one data row.", ex.Message);
            Assert.Equal("items", ex.ParamName);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects_ThrowsWhenColumnsCannotBeInferred() {
            foreach (var items in new[] {
                new object?[] { "A", "B" },
                new object?[] { new object() }
            }) {
                var ex = Assert.Throws<InvalidOperationException>(() => ObjectDataTableBuilder.FromObjects(items));
                Assert.Equal("Unable to infer column names. Use objects with properties or dictionaries.", ex.Message);
            }
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects_ThrowsOnNullFirstRow() {
            var ex = Assert.Throws<ArgumentException>(() => ObjectDataTableBuilder.FromObjects(new object?[] { null }));
            Assert.StartsWith("Data rows cannot be null.", ex.Message);
            Assert.Equal("items", ex.ParamName);
        }

        [Fact]
        public void Test_ObjectDataTableBuilder_FromObjects_ThrowsOnNullEntry() {
            var ex = Assert.Throws<InvalidOperationException>(() => ObjectDataTableBuilder.FromObjects(new object?[] { new { Name = "A" }, null }));
            Assert.Equal("Data rows cannot contain null entries.", ex.Message);
        }

        private class ObjectDataBaseRow {
            public string Name { get; set; } = string.Empty;
            public int Value { get; set; }
        }

        private sealed class InheritedObjectRow : ObjectDataBaseRow {
            public string Notes { get; set; } = string.Empty;
        }

        private sealed class HiddenObjectRow : ObjectDataBaseRow {
            public new string Name { get; set; } = string.Empty;
            public new int Value { get; set; }
        }

        private sealed class EnumerableObjectRow : IEnumerable<int> {
            public string Name { get; set; } = string.Empty;
            public int Value { get; set; }

            public IEnumerator<int> GetEnumerator() {
                yield return Value;
            }

            IEnumerator IEnumerable.GetEnumerator() {
                return GetEnumerator();
            }
        }

    }
}
