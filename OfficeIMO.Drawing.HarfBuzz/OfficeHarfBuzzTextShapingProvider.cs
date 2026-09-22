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

    private OfficeTextShapingResult? ShapeUncached(
        OfficeTextShapingRequest request,
        CachedFontCollection fontCollection,
        Language? language) {
        fontCollection.Shape(request, language, out int glyphCount, out GlyphInfo[] infos, out GlyphPosition[] positions);
        if (glyphCount == 0) return null;
        GC.KeepAlive(request.FontDataForShaping);
        if (infos.Length == 0 || infos.Length != positions.Length) return null;

        IReadOnlyDictionary<int, int> clusterEnds = BuildClusterEnds(infos, request.Text.Length);
        var glyphs = new List<OfficeShapedGlyph>(infos.Length);
        for (int index = 0; index < infos.Length; index++) {
            request.CancellationToken.ThrowIfCancellationRequested();
            int glyphId = checked((int)infos[index].Codepoint);
            int textIndex = checked((int)infos[index].Cluster);
            if (glyphId <= 0 || textIndex < 0 || textIndex >= request.Text.Length) return null;

            int end = clusterEnds[textIndex];
            string unicodeText = request.Text.Substring(textIndex, end - textIndex);
            GlyphPosition position = positions[index];
            glyphs.Add(new OfficeShapedGlyph(
                glyphId,
                unicodeText,
                textIndex,
                position.XAdvance,
                position.YAdvance,
                position.XOffset,
                position.YOffset));
        }

        return new OfficeTextShapingResult(glyphs, request.Direction);
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
        int textLength) {
        int[] starts = infos
            .Select(static info => checked((int)info.Cluster))
            .Where(index => index >= 0 && index < textLength)
            .Distinct()
            .OrderBy(static index => index)
            .ToArray();
        var ends = new Dictionary<int, int>(starts.Length);
        for (int index = 0; index < starts.Length; index++) {
            ends[starts[index]] = index + 1 < starts.Length ? starts[index + 1] : textLength;
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
                buffer.AddUtf16(request.Text);
                buffer.GuessSegmentProperties();
                buffer.Direction = request.Direction switch {
                    OfficeTextDirection.LeftToRight => Direction.LeftToRight,
                    OfficeTextDirection.RightToLeft => Direction.RightToLeft,
                    OfficeTextDirection.TopToBottom => Direction.TopToBottom,
                    _ => buffer.Direction
                };
                if (language != null) {
                    buffer.Language = language;
                }

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
