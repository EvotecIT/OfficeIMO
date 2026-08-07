using System.Collections;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-AotSmoke-" + Guid.NewGuid().ToString("N") + ".pptx");
try {
    using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
        var data = new OfficeChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeChartSeries("Revenue", new[] { 1d, 2d, 3d }) });

        PowerPointSlide slide = presentation.AddSlide();
        slide.AddTitle("OfficeIMO NativeAOT slide");
        slide.AddChart(OfficeChartKind.ColumnClustered, data);
        var dictionaryRows = new[] {
            new GenericOnlyDictionaryRow<int>(("Score", 10), ("Rank", 1)),
            new GenericOnlyDictionaryRow<int>(("Score", 20), ("Rank", 2))
        };
        PowerPointTable dictionaryTable = slide.AddTable(dictionaryRows);
        if (dictionaryTable.GetCell(0, 0).Text != "Score" || dictionaryTable.GetCell(0, 1).Text != "Rank" ||
            dictionaryTable.GetCell(1, 0).Text != "10" || dictionaryTable.GetCell(1, 1).Text != "1") {
            throw new InvalidOperationException("PowerPoint AddTable did not preserve generic-only dictionary columns under NativeAOT.");
        }
        presentation.DuplicateSlide(0);
        presentation.Save();
    }

    using PowerPointPresentation reopened = PowerPointPresentation.Load(path);
    if (reopened.Slides.Count != 2 || reopened.Slides[0].Charts.Count() != 1 || reopened.Slides[1].Charts.Count() != 1 ||
        reopened.Slides[0].Tables.Single().GetCell(1, 0).Text != "10" || reopened.Slides[1].Tables.Single().GetCell(1, 0).Text != "10") {
        throw new InvalidOperationException("The PowerPoint round trip lost its slide or cloned chart relationships.");
    }

    Console.WriteLine("PASS | PowerPoint chart and generic-only dictionary table create, duplicate, save, and reload");
} finally {
    if (File.Exists(path)) File.Delete(path);
}

file sealed class GenericOnlyDictionaryRow<TValue> : IReadOnlyDictionary<string, TValue> {
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
