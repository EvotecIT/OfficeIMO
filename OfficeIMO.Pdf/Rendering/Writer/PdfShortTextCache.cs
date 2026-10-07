using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Bounded reuse of successful default ASCII plans within one document font.
/// Results never include custom providers, diagnostic runs, or explicit features.
/// Callers replay usage and coverage callbacks outside the cache lock.
/// </summary>
internal sealed class PdfShortTextCache<T> where T : class {
    private const int MaximumEntries = 64;
    private const int MaximumGlyphs = 8192;
    private readonly object _lock = new();
    private readonly Dictionary<(string Text, string? Language, OfficeTextDirection Direction), LinkedListNode<Entry>> _entries = new();
    private readonly LinkedList<Entry> _recent = new();
    private int _glyphs;

    internal static bool IsEligible(string text, PdfTextShapingOptions options) {
        if (text.Length == 0 || text.Length > 256 || options.ShapingProvider != null ||
            options.ShapingMode != PdfTextShapingMode.OpenTypeLigatures || !options.FeatureSettings.IsDefault ||
            !options.RecordGlyphUsage || !options.ThrowOnMissingGlyph || options.SkipLayoutControls ||
            options.ReportControlCharacters ||
            (options.Direction != OfficeTextDirection.Auto && options.Direction != OfficeTextDirection.LeftToRight)) return false;
        foreach (char character in text) if (character < 32 || character > 126) return false;
        return true;
    }

    internal bool TryGet(string text, PdfTextShapingOptions options, out T value) {
        lock (_lock) {
            if (_entries.TryGetValue((text, options.Language, options.Direction), out LinkedListNode<Entry>? node)) {
                _recent.Remove(node);
                _recent.AddLast(node);
                value = node.Value.Value;
                return true;
            }
        }
        value = null!;
        return false;
    }

    internal void Add(string text, PdfTextShapingOptions options, T value, int glyphs) {
        if (glyphs > MaximumGlyphs) return;
        lock (_lock) {
            var key = (text, options.Language, options.Direction);
            if (_entries.ContainsKey(key)) return;
            while (_recent.Count >= MaximumEntries || _glyphs + glyphs > MaximumGlyphs) {
                LinkedListNode<Entry> first = _recent.First!;
                _entries.Remove(first.Value.Key);
                _glyphs -= first.Value.Glyphs;
                _recent.RemoveFirst();
            }
            var node = _recent.AddLast(new Entry(key, value, glyphs));
            _entries.Add(key, node);
            _glyphs += glyphs;
        }
    }

    private sealed class Entry {
        internal Entry((string, string?, OfficeTextDirection) key, T value, int glyphs) {
            Key = key; Value = value; Glyphs = glyphs;
        }
        internal (string, string?, OfficeTextDirection) Key { get; }
        internal T Value { get; }
        internal int Glyphs { get; }
    }
}

/// <summary>Nominal measurement and its exact glyph-usage effects, including substitution continuations.</summary>
internal sealed class PdfMeasuredText {
    internal PdfMeasuredText(int advance, List<(int GlyphId, string Unicode)> usage) {
        Advance = advance; Usage = usage.ToArray();
    }
    internal int Advance { get; }
    internal (int GlyphId, string Unicode)[] Usage { get; }
    internal void ReplayUsage(Action<int, string> record) {
        foreach (var glyph in Usage) record(glyph.GlyphId, glyph.Unicode);
    }
}
