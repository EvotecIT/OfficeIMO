using System;
using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData(false, WordFileFormat.Docx)]
        [InlineData(true, WordFileFormat.Docx)]
        [InlineData(false, WordFileFormat.Doc)]
        [InlineData(true, WordFileFormat.Doc)]
        public void ObjectTablesOnlyStyleAndRepeatAnIncludedHeader(bool includeHeader, WordFileFormat format) {
            using WordDocument source = WordDocument.Create();
            var table = source.AddTableFromObjects(new object[] {
                new { Name = "North", Value = 42 }, new { Name = "South", Value = 73 }
            }, includeHeader: includeHeader);
            Assert.Equal(WordTableStyle.OfficeIMOModern, table.Style);
            Assert.Equal(includeHeader, table.ConditionalFormattingFirstRow);
            using WordDocument restored = WordDocument.Load(new MemoryStream(source.ToBytes(new WordSaveOptions {
                Format = format, AllowLoss = true
            })));
            WordTable native = Assert.Single(restored.Tables);
            Assert.Equal(includeHeader ? 3 : 2, native.Rows.Count);
            Assert.Equal(includeHeader ? "Name" : "North", native.Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Equal(includeHeader, native.Rows[0].RepeatHeaderRowAtTheTopOfEachPage);
            if (format == WordFileFormat.Docx) Assert.Equal(includeHeader, native.ConditionalFormattingFirstRow);
            Assert.False(native.Rows[includeHeader ? 1 : 0].Cells[0].Paragraphs[0].Bold);
        }

        [Fact]
        public void Test_AddTableFromObjects_WritesHeadersAndValues() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.FromObjects.docx");

            using (var document = WordDocument.Create(filePath)) {
                var items = new[] {
                    new { Name = "Alpha", Value = 1 },
                    new { Name = "Beta", Value = 2 }
                };

                var table = document.AddTableFromObjects(items, WordTableStyle.TableGrid, includeHeader: true);

                Assert.Equal(3, table.RowsCount);
                Assert.Equal(2, table.Rows[0].CellsCount);
                Assert.Equal("Name", table.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("Value", table.Rows[0].Cells[1].Paragraphs[0].Text);

                Assert.Equal("Alpha", table.Rows[1].Cells[0].Paragraphs[0].Text);
                Assert.Equal("1", table.Rows[1].Cells[1].Paragraphs[0].Text);
                Assert.Equal("Beta", table.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.Equal("2", table.Rows[2].Cells[1].Paragraphs[0].Text);

                document.Save();
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AddTableFromObjects_WithoutHeader() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.FromObjects.NoHeader.docx");

            using (var document = WordDocument.Create(filePath)) {
                var items = new[] {
                    new { Name = "Gamma", Value = 3 }
                };

                var table = document.AddTableFromObjects(items, WordTableStyle.TableGrid, includeHeader: false);

                Assert.Equal(1, table.RowsCount);
                Assert.Equal("Gamma", table.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("3", table.Rows[0].Cells[1].Paragraphs[0].Text);

                document.Save();
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_AddTableFromObjects_ThrowsOnNullItems() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.FromObjects.NullItems.docx");

            try {
                using var document = WordDocument.Create(filePath);
                var ex = Assert.Throws<ArgumentNullException>(() => document.AddTableFromObjects(null!));
                Assert.Equal("items", ex.ParamName);
            } finally {
                File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_AddTableFromObjects_ThrowsOnEmptyItems() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.FromObjects.Empty.docx");

            try {
                using var document = WordDocument.Create(filePath);
                var ex = Assert.Throws<ArgumentException>(() => document.AddTableFromObjects(Array.Empty<object?>()));
                Assert.StartsWith("Provide at least one data row.", ex.Message);
                Assert.Equal("items", ex.ParamName);
            } finally {
                File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_AddTableFromObjects_ThrowsOnNullFirstRow() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.FromObjects.NullFirst.docx");

            try {
                using var document = WordDocument.Create(filePath);
                var ex = Assert.Throws<ArgumentException>(() => document.AddTableFromObjects(new object?[] { null }));
                Assert.StartsWith("Data rows cannot be null.", ex.Message);
                Assert.Equal("items", ex.ParamName);
            } finally {
                File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_AddTableFromObjects_ThrowsOnNullEntry() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.FromObjects.NullEntry.docx");

            try {
                using var document = WordDocument.Create(filePath);
                var ex = Assert.Throws<InvalidOperationException>(() => document.AddTableFromObjects(new object?[] { new { Name = "Alpha", Value = 1 }, null }));
                Assert.Equal("Data rows cannot contain null entries.", ex.Message);
            } finally {
                File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_AddTableFromObjects_ThrowsOnMissingColumns() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.FromObjects.MissingColumns.docx");

            try {
                using var document = WordDocument.Create(filePath);
                var ex = Assert.Throws<InvalidOperationException>(() => document.AddTableFromObjects(new object?[] { new object() }));
                Assert.Equal("Unable to infer column names. Use objects with properties or dictionaries.", ex.Message);
            } finally {
                File.Delete(filePath);
            }
        }
    }
}
