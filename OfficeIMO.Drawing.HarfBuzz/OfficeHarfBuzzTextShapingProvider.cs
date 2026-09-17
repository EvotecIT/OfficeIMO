using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
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
public sealed class OfficeHarfBuzzTextShapingProvider : IOfficeTextShapingProvider {
    private readonly ConditionalWeakTable<object, CachedFontCollection> _fontCache = new();

    /// <summary>Shared provider instance with a weak cache of parsed font faces.</summary>
    public static OfficeHarfBuzzTextShapingProvider Instance { get; } = new();

    private OfficeHarfBuzzTextShapingProvider() {
    }

    /// <inheritdoc />
    public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        request.CancellationToken.ThrowIfCancellationRequested();
        if (request.Text.Length == 0) return null;

        byte[] fontData = request.FontDataForShaping;
        object fontCacheKey = request.FontProgramCacheKeyForShaping ?? fontData;
        CachedFontCollection fontCollection = _fontCache.GetValue(
            fontCacheKey,
            _ => new CachedFontCollection(fontData));

        // Layout re-shapes each run many times (keep-with-next, flow-fit, and draw passes),
        // so the same word arrives here 6+ times per occurrence. Shaping is a pure function of
        // these inputs, so memoize the immutable result per font. Variable-font instances are
        // rare and would bloat the key, so they bypass the cache.
        ShapeCacheKey? cacheKey = request.VariationCoordinatesForShaping.Count == 0
            ? new ShapeCacheKey(request.Text, request.Language, request.FontCollectionIndex ?? 0, request.Direction, request.FeatureSettings)
            : (ShapeCacheKey?)null;
        if (cacheKey.HasValue && fontCollection.ResultCache.TryGetValue(cacheKey.Value, out OfficeTextShapingResult? cachedResult)) {
            return cachedResult;
        }

        OfficeTextShapingResult? result = ShapeUncached(request, fontCollection);
        // Bounded so a long-lived process shaping unbounded distinct text cannot grow without limit;
        // repeated vocabulary (the case that matters) is captured well below the cap.
        if (cacheKey.HasValue && fontCollection.ResultCache.Count < MaxCachedResultsPerFont) {
            fontCollection.ResultCache.TryAdd(cacheKey.Value, result);
        }

        return result;
    }

    private const int MaxCachedResultsPerFont = 8192;

    private static OfficeTextShapingResult? ShapeUncached(OfficeTextShapingRequest request, CachedFontCollection fontCollection) {
        fontCollection.Shape(request, out int glyphCount, out GlyphInfo[] infos, out GlyphPosition[] positions);
        if (glyphCount <= 1) return null;
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
                position.XOffset,
                position.YOffset));
        }

        return new OfficeTextShapingResult(glyphs);
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
        private readonly OfficeTextDirection _direction;
        private readonly OfficeTextFeatureSettings _features;

        internal ShapeCacheKey(string text, string? language, int collectionIndex, OfficeTextDirection direction, OfficeTextFeatureSettings features) {
            _text = text;
            _language = language;
            _collectionIndex = collectionIndex;
            _direction = direction;
            _features = features;
        }

        public bool Equals(ShapeCacheKey other) =>
            _collectionIndex == other._collectionIndex &&
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
                hash = hash * 31 + (int)_direction;
                hash = hash * 31 + _features.GetHashCode();
                return hash;
            }
        }
    }

    private sealed class CachedFontCollection {
        private readonly object _sync = new();
        private readonly Blob _blob;
        private readonly Dictionary<int, CachedFace> _faces = new();
        internal readonly ConcurrentDictionary<ShapeCacheKey, OfficeTextShapingResult?> ResultCache = new();

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

        internal void Shape(
            OfficeTextShapingRequest request,
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
                    _ => buffer.Direction
                };
                if (!string.IsNullOrWhiteSpace(request.Language)) {
                    buffer.Language = new Language(request.Language);
                }

                request.CancellationToken.ThrowIfCancellationRequested();
                cached.Font.Shape(buffer);
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
    }
}
