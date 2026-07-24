using System;
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
    private readonly ConditionalWeakTable<byte[], CachedFontCollection> _fontCache = new();

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
        CachedFontCollection fontCollection = _fontCache.GetValue(
            fontData,
            static data => new CachedFontCollection(data));
        fontCollection.Shape(request, out int glyphCount, out GlyphInfo[] infos, out GlyphPosition[] positions);
        if (glyphCount <= 1) return null;
        GC.KeepAlive(fontData);
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

    private sealed class CachedFontCollection {
        private readonly object _sync = new();
        private readonly Blob _blob;
        private readonly Dictionary<int, CachedFace> _faces = new();

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
