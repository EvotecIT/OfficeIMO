using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed partial class PdfTrueTypeFontProgram {
    // Shaping is a pure function of the immutable font, the text and the shaping options. Its only side
    // effects are glyph-usage records and provider-shaped-text notifications, so a run shaped once (for
    // instance while measuring a line) is reused (when the line is encoded, or the same word comes up again)
    // by replaying exactly those effects. This is shared by every document fork, like the short-text cache.
    private readonly ShapedRunMemo _shapedRuns;

    [ThreadStatic] private static GlyphUsageCapture? _glyphUsageCapture;

    /// <summary>A glyph-usage call recorded while shaping, replayed in order into the program's usage tables.</summary>
    private readonly struct GlyphUsageRecord {
        internal GlyphUsageRecord(int glyphId, int scalar, string? text) {
            GlyphId = glyphId;
            Scalar = scalar;
            Text = text;
        }

        internal int GlyphId { get; }
        internal int Scalar { get; }
        /// <summary>Normalized logical text of a ligature or cluster; null when <see cref="Scalar"/> applies.</summary>
        internal string? Text { get; }
    }

    private sealed class GlyphUsageCapture {
        internal GlyphUsageCapture(PdfTrueTypeFontProgram program, GlyphUsageCapture? previous) {
            Program = program;
            Previous = previous;
        }

        internal PdfTrueTypeFontProgram Program { get; }
        internal GlyphUsageCapture? Previous { get; }
        internal List<GlyphUsageRecord> Records { get; } = new();
        internal List<(string Text, string FontName, bool Arg3, bool Arg4)>? ProviderCalls { get; set; }
    }

    private sealed class ShapedRun {
        internal ShapedRun(PdfGlyphRun run, GlyphUsageRecord[] usage, (string Text, string FontName, bool Arg3, bool Arg4)[] providerCalls) {
            Run = run;
            Usage = usage;
            ProviderCalls = providerCalls;
        }

        internal PdfGlyphRun Run { get; }
        internal GlyphUsageRecord[] Usage { get; }
        internal (string Text, string FontName, bool Arg3, bool Arg4)[] ProviderCalls { get; }
    }

    /// <summary>
    /// Records glyph usage made through <see cref="RecordGlyphUsage(int, int)"/> or
    /// <see cref="RecordGlyphUsage(int, string)"/> on this thread instead of applying it; returns true when captured.
    /// </summary>
    private bool TryCaptureGlyphUsage(int glyphId, int scalar, string? text) {
        GlyphUsageCapture? capture = _glyphUsageCapture;
        if (capture == null || !ReferenceEquals(capture.Program, this)) return false;
        capture.Records.Add(new GlyphUsageRecord(glyphId, scalar, text));
        return true;
    }

    private void ApplyGlyphUsage(IReadOnlyList<GlyphUsageRecord> records) {
        if (records.Count == 0) return;
        lock (_usageLock) {
            for (int index = 0; index < records.Count; index++) {
                GlyphUsageRecord record = records[index];
                if (record.Text != null) RecordNormalizedGlyphUsage(record.GlyphId, record.Text);
                else RecordGlyphUsageHoldingLock(record.GlyphId, record.Scalar);
            }
        }
    }

    private static bool CanMemoizeShaping(string text, PdfTextShapingOptions options) =>
        text.Length > 0 && text.Length <= ShapedRunMemo.MaximumTextLength &&
        // A host provider may observe or vary its own calls; only the built-in shapers are known to be pure.
        (options.ShapingProvider == null || ReferenceEquals(options.ShapingProvider, OfficeManagedTextShapingProvider.Instance));

    /// <summary>
    /// Shapes <paramref name="text"/> with every glyph-usage record applied under one acquisition of the usage
    /// lock, and reuses an identical earlier shaping when one is remembered.
    /// </summary>
    private PdfGlyphRun ShapeAndRecordUsage(string text, PdfTextShapingOptions options, bool allowMemo) {
        bool memoize = allowMemo && CanMemoizeShaping(text, options);
        if (memoize && _shapedRuns.TryGet(text, options, out ShapedRun? remembered)) {
            if (options.RecordGlyphUsage) ApplyGlyphUsage(remembered!.Usage);
            Action<string, string, bool, bool>? recorder = options.ProviderShapedTextRecorder;
            if (recorder != null) {
                foreach ((string Text, string FontName, bool Arg3, bool Arg4) call in remembered!.ProviderCalls) recorder(call.Text, call.FontName, call.Arg3, call.Arg4);
            }
            return remembered!.Run;
        }

        var capture = new GlyphUsageCapture(this, _glyphUsageCapture);
        PdfTextShapingOptions shapingOptions = options;
        if (memoize) {
            Action<string, string, bool, bool>? recorder = options.ProviderShapedTextRecorder;
            shapingOptions = WithProviderShapedTextRecorder(options, (shapedText, fontName, arg3, arg4) => {
                (capture.ProviderCalls ??= new()).Add((shapedText, fontName, arg3, arg4));
                recorder?.Invoke(shapedText, fontName, arg3, arg4);
            });
        }

        PdfGlyphRun result;
        _glyphUsageCapture = capture;
        try {
            result = ShapeUncachedText(text, shapingOptions);
        } finally {
            _glyphUsageCapture = capture.Previous;
            // Usage recorded before a failure (a missing glyph throws) stays recorded, as it did per glyph.
            ApplyGlyphUsage(capture.Records);
        }

        if (memoize && !result.HasMissingGlyphs) {
            _shapedRuns.Add(text, options, new ShapedRun(result, capture.Records.ToArray(),
                capture.ProviderCalls?.ToArray() ?? Array.Empty<(string, string, bool, bool)>()), result.Glyphs.Count);
        }
        return result;
    }

    private static PdfTextShapingOptions WithProviderShapedTextRecorder(PdfTextShapingOptions options,
        Action<string, string, bool, bool> recorder) =>
        new(options.RecordGlyphUsage, options.ThrowOnMissingGlyph, options.SkipLayoutControls, options.ReportControlCharacters,
            options.Source, options.FontName, options.ShapingMode, options.ShapingProvider, recorder, options.Language,
            options.FeatureSettings, options.Direction, options.PreservesPaintedGlyphs);

    /// <summary>
    /// A bounded set of shaped runs keyed by text and every option that shaping reads. Lookups take no lock;
    /// when the set is full it starts over, which keeps the frequent words of the current document.
    /// </summary>
    private sealed class ShapedRunMemo {
        internal const int MaximumTextLength = 1024;
        private const int MaximumEntries = 1024;
        private const int MaximumGlyphs = 32_768;
        private readonly object _lock = new();
        private readonly System.Collections.Concurrent.ConcurrentDictionary<Key, ShapedRun> _entries = new();
        private int _glyphs;

        internal bool TryGet(string text, PdfTextShapingOptions options, out ShapedRun? run) =>
            _entries.TryGetValue(new Key(text, options), out run);

        internal void Add(string text, PdfTextShapingOptions options, ShapedRun run, int glyphs) {
            if (glyphs > MaximumGlyphs) return;
            lock (_lock) {
                if (_entries.Count >= MaximumEntries || _glyphs + glyphs > MaximumGlyphs) {
                    _entries.Clear();
                    _glyphs = 0;
                }
                if (_entries.TryAdd(new Key(text, options), run)) _glyphs += glyphs;
            }
        }

        // Every option shaping reads, copied out of the options so a key never holds the caller's
        // provider-shaped-text recorder (and through it a writer); recorder calls are replayed instead.
        private readonly struct Key : IEquatable<Key> {
            private readonly string _text;
            private readonly string _source;
            private readonly string _fontName;
            private readonly string? _language;
            private readonly OfficeTextFeatureSettings _featureSettings;
            private readonly IOfficeTextShapingProvider? _shapingProvider;
            private readonly PdfTextShapingMode _shapingMode;
            private readonly OfficeTextDirection _direction;
            private readonly int _flags;

            internal Key(string text, PdfTextShapingOptions options) {
                _text = text;
                _source = options.Source;
                _fontName = options.FontName;
                _language = options.Language;
                _featureSettings = options.FeatureSettings;
                _shapingProvider = options.ShapingProvider;
                _shapingMode = options.ShapingMode;
                _direction = options.Direction;
                _flags = (options.RecordGlyphUsage ? 1 : 0) | (options.ThrowOnMissingGlyph ? 2 : 0) |
                    (options.SkipLayoutControls ? 4 : 0) | (options.ReportControlCharacters ? 8 : 0) |
                    (options.PreservesPaintedGlyphs ? 16 : 0);
            }

            public bool Equals(Key other) =>
                _flags == other._flags &&
                _shapingMode == other._shapingMode &&
                _direction == other._direction &&
                ReferenceEquals(_shapingProvider, other._shapingProvider) &&
                string.Equals(_text, other._text, StringComparison.Ordinal) &&
                string.Equals(_source, other._source, StringComparison.Ordinal) &&
                string.Equals(_fontName, other._fontName, StringComparison.Ordinal) &&
                string.Equals(_language, other._language, StringComparison.Ordinal) &&
                _featureSettings.Equals(other._featureSettings);

            public override bool Equals(object? obj) => obj is Key other && Equals(other);

            public override int GetHashCode() {
                unchecked {
                    int hash = StringComparer.Ordinal.GetHashCode(_text);
                    hash = hash * 31 + (int)_shapingMode;
                    hash = hash * 31 + (int)_direction;
                    hash = hash * 31 + _flags;
                    hash = hash * 31 + (_language == null ? 0 : StringComparer.Ordinal.GetHashCode(_language));
                    return hash;
                }
            }
        }
    }
}
