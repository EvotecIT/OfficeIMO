using System;
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

    }
}
