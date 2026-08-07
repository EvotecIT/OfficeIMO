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
    }
}
