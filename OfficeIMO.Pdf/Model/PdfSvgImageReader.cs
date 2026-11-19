using System.Runtime.CompilerServices;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Validates SVG image resources consistently at admission and serialization.</summary>
internal static class PdfSvgImageReader {
    // Parsed drawings by source array. Admission, validation and every placement of the same array then share
    // one parse. Entries live only as long as their array, and a stored copy proves the caller has not changed
    // the array since. Drawings holding embedded raster images are not kept: their pixels depend on the codecs
    // registered when they were decoded.
    private static readonly ConditionalWeakTable<byte[], ParsedSvg> ParsedByArray = new();

    // Parsed drawings by content. Converters read a picture's bytes into a new array for every placement or
    // pass; equal SVG then shares one parse. Bounded, with the same no-raster rule as above.
    private const int MaximumContentEntries = 64;
    private const long MaximumContentBytes = 16L * 1024L * 1024L;
    private static readonly object ContentSync = new();
    private static readonly Dictionary<string, OfficeDrawing> ParsedByContent = new(StringComparer.Ordinal);
    private static long _contentBytes;

    private sealed class ParsedSvg {
        internal ParsedSvg(byte[] source, OfficeDrawing drawing) {
            Source = source;
            Drawing = drawing;
        }

        internal byte[] Source { get; }
        internal OfficeDrawing Drawing { get; }
    }

    /// <summary>Parses SVG bytes, reusing the drawing recorded for this array when its content is unchanged.</summary>
    /// <param name="bytes">SVG bytes.</param>
    /// <param name="remember">Records the parsed drawing for later reads of the same array.</param>
    /// <param name="cancellationToken">Cancels parsing.</param>
    internal static OfficeDrawing Read(byte[] bytes, bool remember = false, CancellationToken cancellationToken = default) {
        Guard.NotNullOrEmpty(bytes, nameof(bytes));
        if (TryGetParsed(bytes, out OfficeDrawing cached)) return cached;
        string contentKey = ComputeContentKey(bytes);
        lock (ContentSync) {
            if (ParsedByContent.TryGetValue(contentKey, out OfficeDrawing? shared)) {
                if (remember) Remember(bytes, shared);
                return shared;
            }
        }

        // Office SVG pictures, such as dense QR codes, stay vector within the budget their stored raster fallback was made with.
        var options = new OfficeSvgDrawingReaderOptions {
            CancellationToken = cancellationToken,
            MaximumGeometryCommands = OfficeIMO.Core.Internal.OfficeSvgPictureFallback.MaximumGeometryCommands
        };
        if (!OfficeSvgDrawingReader.TryRead(bytes, options, out OfficeDrawing? vector, out int unsupported) ||
            vector == null || unsupported != 0) {
            throw new NotSupportedException("OfficeIMO.Pdf cannot import this SVG image completely. Supply an image produced by a renderer supporting its SVG features.");
        }

        if (remember) Remember(bytes, vector);
        if (bytes.LongLength <= MaximumContentBytes / 4 && !ContainsRasterImage(vector)) {
            lock (ContentSync) {
                if (!ParsedByContent.TryGetValue(contentKey, out _)) {
                    // Clear-all eviction keeps this simple, as for prepared raster images.
                    if (ParsedByContent.Count >= MaximumContentEntries || _contentBytes + bytes.LongLength > MaximumContentBytes) {
                        ParsedByContent.Clear();
                        _contentBytes = 0;
                    }
                    ParsedByContent[contentKey] = vector;
                    _contentBytes += bytes.LongLength;
                }
            }
        }
        return vector;
    }

    private static string ComputeContentKey(byte[] bytes) {
#if NET6_0_OR_GREATER
        return Convert.ToBase64String(System.Security.Cryptography.SHA256.HashData(bytes));
#else
        using (var sha256 = System.Security.Cryptography.SHA256.Create()) {
            return Convert.ToBase64String(sha256.ComputeHash(bytes));
        }
#endif
    }

    internal static bool TryGetParsed(byte[] bytes, out OfficeDrawing drawing) {
        if (ParsedByArray.TryGetValue(bytes, out ParsedSvg? cached) && BytesEqual(bytes, cached.Source)) {
            drawing = cached.Drawing;
            return true;
        }
        drawing = null!;
        return false;
    }

    /// <summary>Records the drawing parsed from <paramref name="bytes"/> for later reads of the same array.</summary>
    internal static void Remember(byte[] bytes, OfficeDrawing drawing) {
        if (ContainsRasterImage(drawing)) return;
        var entry = new ParsedSvg((byte[])bytes.Clone(), drawing);
#if NET6_0_OR_GREATER
        ParsedByArray.AddOrUpdate(bytes, entry);
#else
        ParsedByArray.Remove(bytes);
        try {
            ParsedByArray.Add(bytes, entry);
        } catch (ArgumentException) {
            // Another thread recorded the same array first.
        }
#endif
    }

    private static bool ContainsRasterImage(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            switch (element) {
                case OfficeDrawingImage:
                    return true;
                case OfficeDrawingGroup group when ContainsRasterImage(group.InnerDrawing):
                    return true;
                case OfficeDrawingEffectGroup effect when ContainsRasterImage(effect.InnerDrawing) ||
                    effect.SoftMask != null && ContainsRasterImage(effect.SoftMask.InnerDrawing):
                    return true;
                case OfficeDrawingTilingPattern pattern when ContainsRasterImage(pattern.InnerTile):
                    return true;
            }
        }
        return false;
    }

    private static bool BytesEqual(byte[] left, byte[] right) {
        if (ReferenceEquals(left, right)) return true;
#if NET8_0_OR_GREATER
        return left.AsSpan().SequenceEqual(right);
#else
        if (left.Length != right.Length) return false;
        for (int i = 0; i < left.Length; i++) {
            if (left[i] != right[i]) return false;
        }
        return true;
#endif
    }
}
