using System;
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
    }
}
