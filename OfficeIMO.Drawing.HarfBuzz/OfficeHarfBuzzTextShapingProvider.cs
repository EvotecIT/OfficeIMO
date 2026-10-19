using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using HarfBuzzSharp;

namespace OfficeIMO.Drawing.HarfBuzz;

/// <summary>
/// Shapes OfficeIMO text runs with HarfBuzz OpenType GSUB/GPOS processing.
/// </summary>
/// <remarks>
/// The provider is an optional adapter over the shared
/// <see cref="IOfficeTextShapingProvider"/> contract. Core Drawing and PDF
/// packages remain independent of HarfBuzz and its native assets.
/// </remarks>
public sealed class OfficeHarfBuzzTextShapingProvider : IOfficeTextShapingProvider, IOfficeTextShapingProviderMetadata {
    private readonly ConditionalWeakTable<object, CachedFontCollection> _fontCache = new();
    private static readonly object LanguageSync = new();
    private static readonly Dictionary<string, Language> Languages = new(StringComparer.Ordinal);

    /// <summary>Shared provider instance with a weak cache of parsed font faces.</summary>
    public static OfficeHarfBuzzTextShapingProvider Instance { get; } = new();

    internal OfficeHarfBuzzTextShapingProvider() {
    }

    /// <inheritdoc />
    public OfficeTextShapingBackend Backend => OfficeTextShapingBackend.HarfBuzz;

    /// <inheritdoc />
    public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.Text.Length == 0) return null;
        byte[] fontData = request.FontDataForShaping;
        object fontCacheKey = request.FontProgramCacheKeyForShaping ?? fontData;
        ResolvedLanguage language = ResolveLanguage(request.Language);
        if (!_fontCache.TryGetValue(fontCacheKey, out CachedFontCollection? fontCollection)) {
            // Avoid allocating the value-factory closure on the overwhelmingly common cache-hit path.
            fontCollection = _fontCache.GetValue(
                fontCacheKey,
                _ => new CachedFontCollection(fontData));
        }

        // Layout re-shapes each run many times (keep-with-next, flow-fit, and draw passes),
        // so the same word arrives here 6+ times per occurrence. Shaping is a pure function of
        // these inputs, so memoize the immutable result per font. Variable-font instances are
        // rare and would bloat the key, so they bypass the cache.
        ShapeCacheKey? cacheKey = request.VariationCoordinatesForShaping.Count == 0 &&
            request.Text.Length <= MaxCacheableTextLength
            ? new ShapeCacheKey(request.Text, language.CacheKey, request.FontCollectionIndex ?? 0, request.UnitsPerEm, request.Direction, request.FeatureSettings)
            : (ShapeCacheKey?)null;
        if (cacheKey.HasValue && fontCollection.TryGetCachedResult(cacheKey.Value, out OfficeTextShapingResult? cachedResult)) {
            return cachedResult;
        }

        OfficeTextShapingResult? result = ShapeUncached(request, fontCollection, language.HarfBuzzLanguage);
        if (cacheKey.HasValue && result != null) {
            fontCollection.CacheResult(cacheKey.Value, result);
        }

        return result;
    }

    private const int MaxCachedResultsPerFont = 4096;
    private const int MaxCacheableTextLength = 4096;
    private const int MaxCacheableLanguageLength = 255;
    internal const int MaxInternedLanguagesPerProcess = 1024;
    private const long MaxCachedResultBytesPerFont = 8L * 1024L * 1024L;

    // Mixed text is shaped as runs of one script and one bidi level (UAX #9), so each run gets
    // the right OpenType script and direction. Runs are placed in visual order. Each cluster's
    // text belongs to its first glyph; other glyphs of the cluster are extraction continuations.
    private OfficeTextShapingResult? ShapeUncached(
        OfficeTextShapingRequest request,
        CachedFontCollection fontCollection,
        Language? language) {
        string text = request.Text;
        List<ShapingRun> runs = SegmentRuns(text, request.Direction, request.CancellationToken);
        var shaped = new (GlyphInfo[] Infos, GlyphPosition[] Positions)[runs.Count];
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            ShapingRun run = runs[runIndex];
            fontCollection.Shape(request, language, run.Start, run.Length, run.Direction, run.Script,
                out int glyphCount, out GlyphInfo[] infos, out GlyphPosition[] positions);
            if (glyphCount == 0 || infos.Length == 0 || infos.Length != positions.Length) return null;
            shaped[runIndex] = (infos, positions);
        }
        GC.KeepAlive(request.FontDataForShaping);

        int[] visualOrder = runs.Count == 1 ? new[] { 0 }
            : OfficeUnicodeBidi.ReorderItems(runs.Select(static run => run.Level).ToArray());
        var glyphs = new List<OfficeShapedGlyph>();
        var clusterStarts = new List<int>();
        var assigned = new HashSet<int>();
        foreach (int runIndex in visualOrder) {
            ShapingRun run = runs[runIndex];
            (GlyphInfo[] infos, GlyphPosition[] positions) = shaped[runIndex];
            IReadOnlyDictionary<int, int> clusterEnds = BuildClusterEnds(infos, run.Start, run.Start + run.Length);
            for (int index = 0; index < infos.Length; index++) {
                request.CancellationToken.ThrowIfCancellationRequested();
                int glyphId = checked((int)infos[index].Codepoint);
                int textIndex = checked((int)infos[index].Cluster);
                if (glyphId <= 0 || textIndex < run.Start || textIndex >= run.Start + run.Length) return null;
                GlyphPosition position = positions[index];
                if (assigned.Add(textIndex)) {
                    glyphs.Add(new OfficeShapedGlyph(glyphId, text.Substring(textIndex, clusterEnds[textIndex] - textIndex), textIndex,
                        position.XAdvance, position.YAdvance, position.XOffset, position.YOffset));
                } else {
                    glyphs.Add(OfficeShapedGlyph.CreatePositionedUnicodeContinuation(glyphId, textIndex,
                        position.XAdvance, position.YAdvance, position.XOffset, position.YOffset));
                }
                clusterStarts.Add(textIndex);
            }
        }

        OfficeTextDirection direction = request.Direction != OfficeTextDirection.Auto ? request.Direction
            : runs.Count == 1 ? runs[0].Direction : OfficeTextDirection.Auto;
        return new OfficeTextShapingResult(glyphs, advanceAdjustments: null, direction, clusterStarts);
    }

    private readonly struct ShapingRun {
        internal ShapingRun(int start, int length, int level, OfficeTextDirection direction, Script script) {
            Start = start;
            Length = length;
            Level = level;
            Direction = direction;
            Script = script;
        }
        internal int Start { get; }
        internal int Length { get; }
        internal int Level { get; }
        internal OfficeTextDirection Direction { get; }
        internal Script Script { get; }
    }

    private static List<ShapingRun> SegmentRuns(string text, OfficeTextDirection direction, System.Threading.CancellationToken cancellationToken) {
        int[] scalars = OfficeUnicodeBidi.ToScalars(text, out int[] starts);
        if (direction == OfficeTextDirection.TopToBottom) {
            return new List<ShapingRun> { new ShapingRun(0, text.Length, 0, direction, ResolveFirstScript(scalars)) };
        }
        int? paragraphLevel = direction switch {
            OfficeTextDirection.RightToLeft => 1,
            OfficeTextDirection.LeftToRight => 0,
            _ => null
        };
        byte[] levels = OfficeUnicodeBidi.Resolve(scalars, paragraphLevel, cancellationToken).GetLineLevels(0, scalars.Length);
        var scripts = new Script[scalars.Length];
        Script current = Script.Invalid;
        for (int index = 0; index < scalars.Length; index++) {
            Script script = UnicodeFunctions.Default.GetScript(scalars[index]);
            if (script == Script.Common || script == Script.Inherited || script == Script.Unknown || script == Script.Invalid) script = current;
            scripts[index] = script;
            if (script != Script.Invalid) current = script;
        }
        // Leading neutrals take the first real script.
        Script first = ResolveFirstScript(scalars);
        for (int index = 0; index < scripts.Length && scripts[index] == Script.Invalid; index++) scripts[index] = first;

        var runs = new List<ShapingRun>();
        int runStart = 0;
        for (int index = 1; index <= scalars.Length; index++) {
            if (index < scalars.Length && levels[index] == levels[runStart] && scripts[index] == scripts[runStart]) continue;
            int start = starts[runStart];
            int end = index < scalars.Length ? starts[index] : text.Length;
            OfficeTextDirection runDirection = (levels[runStart] & 1) == 1 ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight;
            runs.Add(new ShapingRun(start, end - start, levels[runStart], runDirection, scripts[runStart]));
            runStart = index;
        }
        return runs;
    }

    private static Script ResolveFirstScript(int[] scalars) {
        foreach (int scalar in scalars) {
            Script script = UnicodeFunctions.Default.GetScript(scalar);
            if (script != Script.Common && script != Script.Inherited && script != Script.Unknown && script != Script.Invalid) return script;
        }
        return Script.Common;
    }

    private ResolvedLanguage ResolveLanguage(string? value) {
        string? normalized = NormalizeLanguage(value);
        if (normalized == null) {
            return default;
        }

        lock (LanguageSync) {
            if (Languages.TryGetValue(normalized, out Language? language)) {
                return new ResolvedLanguage(normalized, language);
            }
            if (Languages.Count >= MaxInternedLanguagesPerProcess) {
                // HarfBuzz interns language strings for the process lifetime. Shape with
                // its inferred language after saturation so one document cannot make
                // later unrelated requests fail or grow the native intern table.
                return default;
            }

            language = new Language(normalized);
            Languages.Add(normalized, language);
            return new ResolvedLanguage(normalized, language);
        }
    }

    private static string? NormalizeLanguage(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string language = value!.Trim();
        if (language.Length == 0 || language.Length > MaxCacheableLanguageLength) {
            return null;
        }

        int subtagLength = 0;
        for (int index = 0; index < language.Length; index++) {
            char character = language[index];
            if (character == '-') {
                if (subtagLength == 0 || subtagLength > 8) {
                    return null;
                }
                subtagLength = 0;
                continue;
            }
            if (!((character >= 'A' && character <= 'Z') ||
                  (character >= 'a' && character <= 'z') ||
                  (character >= '0' && character <= '9'))) {
                return null;
            }
            subtagLength++;
        }
        if (subtagLength == 0 || subtagLength > 8) {
            return null;
        }

        return language.ToLowerInvariant();
    }

    private static IReadOnlyDictionary<int, int> BuildClusterEnds(
        IReadOnlyList<GlyphInfo> infos,
        int runStart,
        int runEnd) {
        int[] starts = infos
            .Select(static info => checked((int)info.Cluster))
            .Where(index => index >= runStart && index < runEnd)
            .Distinct()
            .OrderBy(static index => index)
            .ToArray();
        var ends = new Dictionary<int, int>(starts.Length);
        for (int index = 0; index < starts.Length; index++) {
            ends[starts[index]] = index + 1 < starts.Length ? starts[index + 1] : runEnd;
        }
        return ends;
    }

    private readonly struct ShapeCacheKey : IEquatable<ShapeCacheKey> {
        private readonly string _text;
        private readonly string? _language;
        private readonly int _collectionIndex;
        private readonly int _unitsPerEm;
        private readonly OfficeTextDirection _direction;
        private readonly OfficeTextFeatureSettings _features;

        internal ShapeCacheKey(string text, string? language, int collectionIndex, int unitsPerEm, OfficeTextDirection direction, OfficeTextFeatureSettings features) {
            _text = text;
            _language = language;
            _collectionIndex = collectionIndex;
            _unitsPerEm = unitsPerEm;
            _direction = direction;
            _features = features;
        }

        public bool Equals(ShapeCacheKey other) =>
            _collectionIndex == other._collectionIndex &&
            _unitsPerEm == other._unitsPerEm &&
            _direction == other._direction &&
            string.Equals(_text, other._text, StringComparison.Ordinal) &&
            string.Equals(_language, other._language, StringComparison.Ordinal) &&
            _features.Equals(other._features);

        public override bool Equals(object? obj) => obj is ShapeCacheKey other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = 17;
                hash = hash * 31 + StringComparer.Ordinal.GetHashCode(_text);
                hash = hash * 31 + (_language != null ? StringComparer.Ordinal.GetHashCode(_language) : 0);
                hash = hash * 31 + _collectionIndex;
                hash = hash * 31 + _unitsPerEm;
                hash = hash * 31 + (int)_direction;
                hash = hash * 31 + (_features.IsDefault ? 0 : _features.GetHashCode());
                return hash;
            }
        }

        internal int TextLength => _text.Length;

        internal int LanguageLength => _language?.Length ?? 0;

        internal int FeatureCount => _features.Features.Count;
    }

    private readonly struct ResolvedLanguage {
        internal ResolvedLanguage(string cacheKey, Language harfBuzzLanguage) {
            CacheKey = cacheKey;
            HarfBuzzLanguage = harfBuzzLanguage;
        }

        internal string? CacheKey { get; }

        internal Language? HarfBuzzLanguage { get; }
    }

    private sealed class CachedFontCollection {
        private readonly object _sync = new();
        private readonly object _resultCacheSync = new();
        private readonly Blob _blob;
        private readonly Dictionary<int, CachedFace> _faces = new();
        private readonly Dictionary<ShapeCacheKey, LinkedListNode<CachedShapeResult>> _resultCache = new();
        private readonly LinkedList<CachedShapeResult> _resultCacheLru = new();
        private long _resultCacheBytes;

        internal CachedFontCollection(byte[] fontData) {
            GCHandle pinned = GCHandle.Alloc(fontData, GCHandleType.Pinned);
            try {
                _blob = new Blob(
                    pinned.AddrOfPinnedObject(),
                    fontData.Length,
                    MemoryMode.Duplicate);
            } finally {
                pinned.Free();
            }
        }

        internal bool TryGetCachedResult(ShapeCacheKey key, out OfficeTextShapingResult? result) {
            lock (_resultCacheSync) {
                if (_resultCache.TryGetValue(key, out LinkedListNode<CachedShapeResult>? node)) {
                    _resultCacheLru.Remove(node);
                    _resultCacheLru.AddFirst(node);
                    result = node.Value.Result;
                    return true;
                }
            }

            result = null;
            return false;
        }

        internal void CacheResult(ShapeCacheKey key, OfficeTextShapingResult result) {
            long estimatedBytes = EstimateCacheSizeBytes(key, result);
            if (estimatedBytes > MaxCachedResultBytesPerFont) {
                return;
            }

            lock (_resultCacheSync) {
                if (_resultCache.TryGetValue(key, out LinkedListNode<CachedShapeResult>? existing)) {
                    _resultCacheLru.Remove(existing);
                    _resultCacheLru.AddFirst(existing);
                    return;
                }

                var entry = new CachedShapeResult(key, result, estimatedBytes);
                LinkedListNode<CachedShapeResult> node = _resultCacheLru.AddFirst(entry);
                _resultCache.Add(key, node);
                _resultCacheBytes += estimatedBytes;

                while (_resultCache.Count > MaxCachedResultsPerFont ||
                       _resultCacheBytes > MaxCachedResultBytesPerFont) {
                    LinkedListNode<CachedShapeResult>? last = _resultCacheLru.Last;
                    if (last == null) {
                        break;
                    }

                    _resultCacheLru.RemoveLast();
                    _resultCache.Remove(last.Value.Key);
                    _resultCacheBytes -= last.Value.EstimatedBytes;
                }
            }
        }

        private static long EstimateCacheSizeBytes(ShapeCacheKey key, OfficeTextShapingResult result) {
            long bytes = 128L +
                (key.TextLength * 2L) +
                (key.LanguageLength * 2L) +
                (key.FeatureCount * 72L) +
                (result.Glyphs.Count * 48L);
            for (int index = 0; index < result.Glyphs.Count; index++) {
                bytes += result.Glyphs[index].UnicodeText.Length * 2L;
            }

            return bytes;
        }

        internal void Shape(
            OfficeTextShapingRequest request,
            Language? language,
            int start,
            int length,
            OfficeTextDirection direction,
            Script script,
            out int glyphCount,
            out GlyphInfo[] infos,
            out GlyphPosition[] positions) {
            lock (_sync) {
                int collectionIndex = request.FontCollectionIndex ?? 0;
                if (!_faces.TryGetValue(collectionIndex, out CachedFace? cached)) {
                    var face = new Face(_blob, collectionIndex);
                    var font = new Font(face);
                    font.SetFunctionsOpenType();
                    cached = new CachedFace(face, font);
                    _faces.Add(collectionIndex, cached);
                }

                glyphCount = cached.Face.GlyphCount;
                cached.Font.SetScale(request.UnitsPerEm, request.UnitsPerEm);
                Variation[] variations = request.VariationCoordinatesForShaping
                    .Select(static coordinate => new Variation {
                        Tag = HarfBuzzSharp.Tag.Parse(coordinate.Key),
                        Value = coordinate.Value
                    })
                    .ToArray();
                cached.Font.SetVariations(variations);
                using var buffer = new HarfBuzzSharp.Buffer();
                // The whole text is context for this run; clusters are indexes into it.
                buffer.AddUtf16(request.Text, start, length);
                buffer.Direction = direction switch {
                    OfficeTextDirection.RightToLeft => Direction.RightToLeft,
                    OfficeTextDirection.TopToBottom => Direction.TopToBottom,
                    _ => Direction.LeftToRight
                };
                buffer.Script = script;
                if (language != null) {
                    buffer.Language = language;
                }
                buffer.GuessSegmentProperties();

                request.CancellationToken.ThrowIfCancellationRequested();
                Feature[] features = request.FeatureSettings.Features
                    .Select(static feature => new Feature(
                        HarfBuzzSharp.Tag.Parse(feature.Key),
                        checked((uint)feature.Value)))
                    .ToArray();
                cached.Font.Shape(buffer, features);
                infos = buffer.GlyphInfos;
                positions = buffer.GlyphPositions;
            }
        }

        ~CachedFontCollection() {
            foreach (CachedFace cached in _faces.Values) {
                cached.Font.Dispose();
                cached.Face.Dispose();
            }
            _blob.Dispose();
        }

        private sealed record CachedFace(Face Face, Font Font);

        private sealed record CachedShapeResult(
            ShapeCacheKey Key,
            OfficeTextShapingResult Result,
            long EstimatedBytes);
    }
}
