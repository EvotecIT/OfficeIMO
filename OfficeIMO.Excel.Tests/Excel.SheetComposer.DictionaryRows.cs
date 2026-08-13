using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Data;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelSheetComposerDictionaryRowTests {
        [Fact]
        public void ObjectFlattener_SupportsDictionaryRowsAndNestedCellValues() {
            var metadata = new ReadOnlyDictionaryRow(
                ("Region", "EU"),
                ("Tier", 1));
            var row = new ReadOnlyDictionaryRow(
                ("Name", "Alpha"),
                ("Tags", new[] { "One", "Two" }),
                ("Metadata", metadata));
            var options = new ObjectFlattenerOptions {
                CollectionJoinWith = " | ",
                DictionaryEntryJoinWith = "; ",
                DictionaryKeyValueSeparator = ": "
            };

            var flattened = new ObjectFlattener().Flatten(row, options);

            Assert.Equal("Alpha", flattened["Name"]);
            Assert.Equal("One | Two", flattened["Tags"]);
            Assert.Equal("Region: EU; Tier: 1", flattened["Metadata"]);
        }

        [Fact]
        public void ObjectFlattener_SupportsGenericOnlyMutableDictionaryRows() {
            var row = new GenericOnlyMutableDictionaryRow<int>(("Score", 10), ("Rank", 1));

            Dictionary<string, object?> flattened = new ObjectFlattener().Flatten(
                row,
                new ObjectFlattenerOptions());

            Assert.Equal(10, flattened["Score"]);
            Assert.Equal(1, flattened["Rank"]);
            Assert.DoesNotContain("Keys", flattened.Keys);
            Assert.DoesNotContain("Values", flattened.Keys);
            Assert.DoesNotContain("Count", flattened.Keys);
        }

        [Fact]
        public void ObjectFlattener_EnforcesDictionaryItemLimitAcrossAdapters() {
            var options = new ObjectFlattenerOptions { MaxCollectionItems = 1 };
            var nonGeneric = new Hashtable { ["Score"] = 10, ["Rank"] = 1 };
            var genericOnly = new GenericOnlyMutableDictionaryRow<int>(("Score", 10), ("Rank", 1));
            var readOnly = new GenericOnlyDictionaryRow<int>(("Score", 10), ("Rank", 1));
            var flattener = new ObjectFlattener();

            Assert.Throws<InvalidDataException>(() => flattener.Flatten(nonGeneric, options));
            Assert.Throws<InvalidDataException>(() => flattener.Flatten(genericOnly, options));
            Assert.Throws<InvalidDataException>(() => flattener.Flatten(readOnly, options));
        }

        [Fact]
        public void ObjectFlattener_EnforcesDictionaryColumnLimitBeforeProjection() {
            var row = new ReadOnlyDictionaryRow(("A", 1), ("B", 2), ("C", 3));
            var options = new ObjectFlattenerOptions {
                MaxColumns = 2,
                MaxCollectionItems = 100
            };

            Assert.Throws<InvalidDataException>(() =>
                new ObjectFlattener().Flatten(row, options));
        }

        [Fact]
        public void SheetComposer_EnforcesProjectedCellLimitBeforeWritingCells() {
            using var stream = new MemoryStream();
            using ExcelDocument document = ExcelDocument.Create(stream);
            var rows = new[] {
                new ReadOnlyDictionaryRow(("A", 1), ("B", 2)),
                new ReadOnlyDictionaryRow(("A", 3), ("B", 4))
            };

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => document.Compose("Report", composer =>
                composer.TableFrom(rows, configure: options => options.MaxCells = 5)));

            Assert.Contains("requires at least 6 cells", exception.Message, StringComparison.Ordinal);
            Assert.Contains("2 data rows + 1 header row x 2 columns", exception.Message, StringComparison.Ordinal);
            Assert.Contains("options.MaxCells = 6", exception.Message, StringComparison.Ordinal);
            Assert.Contains("TableFrom(DataTable)", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void SheetComposer_ExplainsHowToOverrideRowAndColumnLimits() {
            var rows = new[] {
                new ReadOnlyDictionaryRow(("A", 1), ("B", 2)),
                new ReadOnlyDictionaryRow(("A", 3), ("B", 4))
            };

            using var rowDocument = ExcelDocument.Create();
            InvalidDataException rowException = Assert.Throws<InvalidDataException>(() =>
                rowDocument.Compose("Rows", composer =>
                    composer.TableFrom(rows, configure: options => options.MaxRows = 1)));
            Assert.Contains("requires at least 2 data rows", rowException.Message, StringComparison.Ordinal);
            Assert.Contains("options.MaxRows = 2", rowException.Message, StringComparison.Ordinal);

            using var columnDocument = ExcelDocument.Create();
            InvalidDataException columnException = Assert.Throws<InvalidDataException>(() =>
                columnDocument.Compose("Columns", composer =>
                    composer.TableFrom(rows, configure: options => options.MaxColumns = 1)));
            Assert.Contains("requires at least 2 columns", columnException.Message, StringComparison.Ordinal);
            Assert.Contains("options.MaxColumns = 2", columnException.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void SheetComposer_ExplainsHardExcelColumnBoundaryForObjectRows() {
            (string Key, object? Value)[] entries = Enumerable.Range(0, A1.MaxColumns + 1)
                .Select(index => ("Column" + index, (object?)index))
                .ToArray();
            var rows = new[] { new ReadOnlyDictionaryRow(entries) };
            using ExcelDocument document = ExcelDocument.Create();

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                document.Compose("Data", composer => composer.TableFrom(
                    rows,
                    configure: options => options.MaxColumns = int.MaxValue,
                    freezeHeaderRow: false)));

            Assert.Contains("requires at least 16385 columns", exception.Message, StringComparison.Ordinal);
            Assert.Contains("Select fewer columns or split the data across multiple worksheets", exception.Message, StringComparison.Ordinal);
            Assert.Contains("cannot be overridden", exception.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("options.MaxColumns = 16385", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void SheetComposer_RendersDictionaryRowsAsRealColumns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                var rows = new[] {
                    new ReadOnlyDictionaryRow(("Name", "Alpha"), ("Score", 10)),
                    new ReadOnlyDictionaryRow(("Name", "Beta"), ("Score", 20))
                };

                using (var document = ExcelDocument.Create(filePath)) {
                    document.Compose("Report", composer => {
                        composer.TableFrom(rows, title: "Scores");
                        composer.Finish(autoFitColumns: false);
                    });
                    document.Save();
                }

                using (var document = ExcelDocument.Load(
                    filePath,
                    new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                    var sheet = document["Report"];
                    Assert.True(sheet.TryGetCellText(2, 1, out string? nameHeader));
                    Assert.True(sheet.TryGetCellText(2, 2, out string? scoreHeader));
                    Assert.True(sheet.TryGetCellText(3, 1, out string? name));
                    Assert.True(sheet.TryGetCellText(3, 2, out string? score));
                    Assert.Equal("Name", nameHeader);
                    Assert.Equal("Score", scoreHeader);
                    Assert.Equal("Alpha", name);
                    Assert.Equal("10", score);
                }
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void SheetBuilder_RendersGenericOnlyDictionaryRowsAsRealColumns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                var rows = new[] {
                    new GenericOnlyDictionaryRow<int>(("Score", 10), ("Rank", 1)),
                    new GenericOnlyDictionaryRow<int>(("Score", 20), ("Rank", 2))
                };

                using (var document = ExcelDocument.Create(filePath)) {
                    document.AsFluent()
                        .Sheet("Report", sheet => sheet.RowsFrom(rows, options =>
                            options.Columns = new[] { "Rank", "Score", "Missing" }))
                        .End()
                        .Save();
                }

                using (var document = ExcelDocument.Load(
                    filePath,
                    new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                    var sheet = document["Report"];
                    Assert.True(sheet.TryGetCellText(1, 1, out string? rankHeader));
                    Assert.True(sheet.TryGetCellText(1, 2, out string? scoreHeader));
                    Assert.True(sheet.TryGetCellText(1, 3, out string? missingHeader));
                    Assert.True(sheet.TryGetCellText(2, 1, out string? rank));
                    Assert.True(sheet.TryGetCellText(2, 2, out string? score));
                    Assert.Equal("Rank", rankHeader);
                    Assert.Equal("Score", scoreHeader);
                    Assert.Equal("Missing", missingHeader);
                    Assert.Equal("1", rank);
                    Assert.Equal("10", score);
                }
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        private sealed class ReadOnlyDictionaryRow : IReadOnlyDictionary<string, object?> {
            private readonly KeyValuePair<string, object?>[] _entries;
            private readonly Dictionary<string, object?> _lookup;

            internal ReadOnlyDictionaryRow(params (string Key, object? Value)[] entries) {
                _entries = entries.Select(entry => new KeyValuePair<string, object?>(entry.Key, entry.Value)).ToArray();
                _lookup = _entries.ToDictionary(entry => entry.Key, entry => entry.Value, StringComparer.OrdinalIgnoreCase);
            }

            public object? this[string key] => _lookup[key];
            public IEnumerable<string> Keys => _entries.Select(entry => entry.Key);
            public IEnumerable<object?> Values => _entries.Select(entry => entry.Value);
            public int Count => _entries.Length;
            public bool ContainsKey(string key) => _lookup.ContainsKey(key);
            public bool TryGetValue(string key, out object? value) => _lookup.TryGetValue(key, out value);
            public IEnumerator<KeyValuePair<string, object?>> GetEnumerator() => ((IEnumerable<KeyValuePair<string, object?>>)_entries).GetEnumerator();
            IEnumerator IEnumerable.GetEnumerator() => _entries.GetEnumerator();
        }

        private sealed class GenericOnlyDictionaryRow<TValue> : IReadOnlyDictionary<string, TValue> {
            private readonly KeyValuePair<string, TValue>[] _entries;
            private readonly Dictionary<string, TValue> _lookup;

            internal GenericOnlyDictionaryRow(params (string Key, TValue Value)[] entries) {
                _entries = entries.Select(entry => new KeyValuePair<string, TValue>(entry.Key, entry.Value)).ToArray();
                _lookup = _entries.ToDictionary(entry => entry.Key, entry => entry.Value, StringComparer.OrdinalIgnoreCase);
            }

            public TValue this[string key] => _lookup[key];
            public IEnumerable<string> Keys => _entries.Select(entry => entry.Key);
            public IEnumerable<TValue> Values => _entries.Select(entry => entry.Value);
            public int Count => _entries.Length;
            public bool ContainsKey(string key) => _lookup.ContainsKey(key);
            public bool TryGetValue(string key, out TValue value) => _lookup.TryGetValue(key, out value!);
            public IEnumerator<KeyValuePair<string, TValue>> GetEnumerator() => ((IEnumerable<KeyValuePair<string, TValue>>)_entries).GetEnumerator();
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }

        private sealed class GenericOnlyMutableDictionaryRow<TValue> : IDictionary<string, TValue> {
            private readonly Dictionary<string, TValue> _values;

            internal GenericOnlyMutableDictionaryRow(params (string Key, TValue Value)[] entries) {
                _values = entries.ToDictionary(entry => entry.Key, entry => entry.Value, StringComparer.OrdinalIgnoreCase);
            }

            public TValue this[string key] { get => _values[key]; set => _values[key] = value; }
            public ICollection<string> Keys => _values.Keys;
            public ICollection<TValue> Values => _values.Values;
            public int Count => _values.Count;
            public bool IsReadOnly => false;
            public void Add(string key, TValue value) => _values.Add(key, value);
            public void Add(KeyValuePair<string, TValue> item) => ((ICollection<KeyValuePair<string, TValue>>)_values).Add(item);
            public void Clear() => _values.Clear();
            public bool Contains(KeyValuePair<string, TValue> item) => ((ICollection<KeyValuePair<string, TValue>>)_values).Contains(item);
            public bool ContainsKey(string key) => _values.ContainsKey(key);
            public void CopyTo(KeyValuePair<string, TValue>[] array, int arrayIndex) => ((ICollection<KeyValuePair<string, TValue>>)_values).CopyTo(array, arrayIndex);
            public IEnumerator<KeyValuePair<string, TValue>> GetEnumerator() => _values.GetEnumerator();
            public bool Remove(string key) => _values.Remove(key);
            public bool Remove(KeyValuePair<string, TValue> item) => ((ICollection<KeyValuePair<string, TValue>>)_values).Remove(item);
            public bool TryGetValue(string key, out TValue value) => _values.TryGetValue(key, out value!);
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}
