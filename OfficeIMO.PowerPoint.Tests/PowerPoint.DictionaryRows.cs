using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointDictionaryRowTests {
        [Fact]
        public void AddTable_RendersGenericOnlyDictionaryRowsAsRealColumns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                var rows = new[] {
                    new GenericOnlyDictionaryRow<int>(("Score", 10), ("Rank", 1)),
                    new GenericOnlyDictionaryRow<int>(("Score", 20), ("Rank", 2))
                };

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(rows);
                    Assert.Equal("Score", table.GetCell(0, 0).Text);
                    Assert.Equal("Rank", table.GetCell(0, 1).Text);
                    Assert.Equal("10", table.GetCell(1, 0).Text);
                    Assert.Equal("1", table.GetCell(1, 1).Text);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointTable table = presentation.Slides.Single().Tables.Single();
                    Assert.Equal("Score", table.GetCell(0, 0).Text);
                    Assert.Equal("Rank", table.GetCell(0, 1).Text);
                    Assert.Equal("10", table.GetCell(1, 0).Text);
                    Assert.Equal("1", table.GetCell(1, 1).Text);
                }
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
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
    }
}
