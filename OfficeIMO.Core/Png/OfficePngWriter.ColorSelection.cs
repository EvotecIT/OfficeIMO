using System.IO.Compression;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    // Palette entries retain straight alpha and hidden RGB. First occurrence order is
    // stable; never quantize, reorder by a randomized dictionary enumeration, or merge
    // transparent colors. Candidate sizes include PLTE/tRNS chunk overhead.
    private sealed class PngColorLayout {
        internal PngColorLayout(int type, int depth, Dictionary<uint, int>? palette = null) {
            Type = type; Depth = depth; Palette = palette;
            if (palette == null) return;
            Colors = new uint[palette.Count];
            foreach (var item in palette) Colors[item.Value] = item.Key;
            for (int i = 0; i < Colors.Length; i++)
                if ((byte)Colors[i] != 255) AlphaCount = i + 1;
        }
        internal int Type { get; }
        internal int Depth { get; }
        internal Dictionary<uint, int>? Palette { get; }
        internal uint[]? Colors { get; }
        internal int AlphaCount { get; }
        internal int Channels => Type == 0 ? 1 : Type == 4 ? 2 : Type == 2 ? 3 : Type == 3 ? 1 : 4;
        internal int RowBytes(int width) => checked((int)(((long)width * Channels * Depth + 7L) / 8L));
        internal int FilterBytesPerPixel => Math.Max(1, Channels * Depth / 8);
        internal long MetadataBytes => Colors == null ? 0 : Colors.Length * 3L + 12 + (AlphaCount == 0 ? 0 : AlphaCount + 12);
    }

    private const int GrayDepthOne = 1, GrayDepthTwo = 2, GrayDepthFour = 4;
    // The gray depths each 8-bit sample can be stored at without loss.
    private static readonly byte[] GrayDepthTable = CreateGrayDepthTable();

    private static byte[] CreateGrayDepthTable() {
        var table = new byte[256];
        for (int value = 0; value < table.Length; value++) {
            table[value] = (byte)((value == 0 || value == 255 ? GrayDepthOne : 0) |
                (value % 85 == 0 ? GrayDepthTwo : 0) | (value % 17 == 0 ? GrayDepthFour : 0));
        }
        return table;
    }

    private static bool IsOpaque(byte[] rgba, int start, int end) {
        int index = start;
#if NET8_0_OR_GREATER
        ReadOnlySpan<uint> pixels = System.Runtime.InteropServices.MemoryMarshal.Cast<byte, uint>(rgba.AsSpan(start, end - start));
        var alpha = BitConverter.IsLittleEndian ? new System.Numerics.Vector<uint>(0xFF000000U) : new System.Numerics.Vector<uint>(0x000000FFU);
        int vectorEnd = pixels.Length - System.Numerics.Vector<uint>.Count, offset = 0;
        for (; offset <= vectorEnd; offset += System.Numerics.Vector<uint>.Count) {
            if (!System.Numerics.Vector.EqualsAll(new System.Numerics.Vector<uint>(pixels.Slice(offset)) & alpha, alpha)) return false;
        }
        index = start + offset * 4;
#endif
        for (; index < end; index += 4) if (rgba[index + 3] != 255) return false;
        return true;
    }

    private static uint RgbaKey(byte[] pixels, int i) =>
        (uint)pixels[i] << 24 | (uint)pixels[i + 1] << 16 | (uint)pixels[i + 2] << 8 | pixels[i + 3];

    private static void EncodeAdaptiveRgba(int width, int height, byte[] rgba, Stream output,
        double? dpiX, double? dpiY, CancellationToken token,
        Action<OfficeRasterEncodingCheckpoint>? observer, bool requireRgbColorSpace, bool requireGrayColorSpace, OfficePngCompression compression, OfficeImageEncodingEffort effort) {
        bool gray = true, opaque = true;
        int grayDepths = GrayDepthOne | GrayDepthTwo | GrayDepthFour;
        Dictionary<uint, int>? palette = new Dictionary<uint, int>();
        uint lastKey = 0;
        bool hasLastKey = false;
        // Blocks keep the checkpoint after every 4096 bytes. Repeated colors skip the palette
        // lookup, and once only opacity is undecided the remaining pixels are checked in bulk.
        for (int block = 0; block < rgba.Length; block += 4096) {
            observer?.Invoke(OfficeRasterEncodingCheckpoint.PngColorSelectionBlock);
            token.ThrowIfCancellationRequested();
            int end = Math.Min(rgba.Length, block + 4096);
            if (palette == null && !gray) {
                if (opaque) opaque = IsOpaque(rgba, block, end);
                if (!opaque && observer == null) break;
                continue;
            }
            for (int i = block; i < end; i += 4) {
                byte g = rgba[i];
                if (gray) {
                    if (g == rgba[i + 1] && g == rgba[i + 2]) grayDepths &= GrayDepthTable[g];
                    else gray = false;
                }
                opaque &= rgba[i + 3] == 255;
                if (palette != null) {
                    uint key = RgbaKey(rgba, i);
                    if (hasLastKey && key == lastKey) continue;
                    if (!palette.ContainsKey(key)) {
                        if (palette.Count == 256) palette = null;
                        else palette.Add(key, palette.Count);
                    }
                    lastKey = key;
                    hasLastKey = true;
                }
            }
        }
        // Bit depths only matter for opaque gray pixels, so they need not be tracked once color appears.
        bool one = (grayDepths & GrayDepthOne) != 0, two = (grayDepths & GrayDepthTwo) != 0, four = (grayDepths & GrayDepthFour) != 0;
        if (requireGrayColorSpace && !gray) throw new ArgumentException("A grayscale ICC profile requires grayscale source pixels.", nameof(rgba));
        if (requireRgbColorSpace) gray = false;
        var layout = new PngColorLayout(gray ? opaque ? 0 : 4 : opaque ? 2 : 6,
            gray && opaque ? one ? 1 : two ? 2 : four ? 4 : 8 : 8);
        if (compression != OfficePngCompression.Optimal) {
            WriteColorPng(output, width, height, rgba, null, layout, false, dpiX, dpiY, token, observer,
                level: compression == OfficePngCompression.Fastest ? CompressionLevel.Fastest : CompressionLevel.NoCompression);
            return;
        }
        if (effort != OfficeImageEncodingEffort.Smallest) {
            // Select by packed sample storage, including palette overhead. This is a
            // cheap size estimate, not a second compression of the same pixels.
            if (palette != null && !requireGrayColorSpace) {
                var indexed = new PngColorLayout(3, palette.Count <= 2 ? 1 : palette.Count <= 4 ? 2 : palette.Count <= 16 ? 4 : 8, palette);
                if ((long)indexed.RowBytes(width) * height + indexed.MetadataBytes < (long)layout.RowBytes(width) * height)
                    layout = indexed;
            }
            WriteColorPng(output, width, height, rgba, null, layout, true, dpiX, dpiY, token, observer,
                level: effort == OfficeImageEncodingEffort.Fast ? CompressionLevel.Fastest : CompressionLevel.Optimal,
                sampledFilters: true);
            return;
        }
        // Compare actual deflated output rather than choosing a palette merely
        // because it has fewer samples: palette overhead dominates small images.
        EnsureColorWorkingSet(rgba.LongLength, layout.RowBytes(width));
        var workspace = new PngColorWorkspace(layout.RowBytes(width));
        var best = ProbeColor(width, height, rgba, layout, workspace, token, observer);
        if (palette != null && !requireGrayColorSpace) {
            var indexed = new PngColorLayout(3, palette.Count <= 2 ? 1 : palette.Count <= 4 ? 2 : palette.Count <= 16 ? 4 : 8, palette);
            var candidate = ProbeColor(width, height, rgba, indexed, workspace, token, observer);
            if (ColorPayloadBytes(candidate.Size) + indexed.MetadataBytes < ColorPayloadBytes(best.Size) + layout.MetadataBytes) {
                layout = indexed; best = candidate;
            }
        }
        WriteColorPng(output, width, height, rgba, null, layout, best.Adaptive,
            dpiX, dpiY, token, observer, workspace);
    }

    private static (long Size, bool Adaptive) ProbeColor(int width, int height, byte[] rgba,
        PngColorLayout layout, PngColorWorkspace workspace, CancellationToken token,
        Action<OfficeRasterEncodingCheckpoint>? observer) {
        using var adaptive = new PngSizeProbeStream();
        WriteColorZlib(adaptive, width, height, rgba, null, layout, true, token, observer, workspace: workspace);
        using var none = new PngSizeProbeStream();
        WriteColorZlib(none, width, height, rgba, null, layout, false, token, observer, workspace: workspace);
        return adaptive.Length <= none.Length ? (adaptive.Length, true) : (none.Length, false);
    }

    private static void WriteColorPng(Stream output, int width, int height, byte[]? rgba,
        ushort[]? rgba64, PngColorLayout layout, bool adaptive, double? dpiX, double? dpiY,
        CancellationToken token, Action<OfficeRasterEncodingCheckpoint>? observer,
        PngColorWorkspace? workspace = null, CompressionLevel? level = null, bool sampledFilters = false) {
        token.ThrowIfCancellationRequested();
        output.Write(PngSignature, 0, PngSignature.Length);
        WriteChunk(output, "IHDR", BuildIhdr(width, height, layout.Depth, layout.Type));
        WriteSrgbChunks(output);
        if (dpiX.HasValue && dpiY.HasValue) WriteChunk(output, "pHYs", BuildPhysicalResolution(dpiX.Value, dpiY.Value));
        if (layout.Colors != null) {
            var colors = new byte[layout.Colors.Length * 3];
            var alpha = new byte[layout.AlphaCount];
            for (int i = 0; i < layout.Colors.Length; i++) {
                uint c = layout.Colors[i]; colors[i * 3] = (byte)(c >> 24);
                colors[i * 3 + 1] = (byte)(c >> 16); colors[i * 3 + 2] = (byte)(c >> 8);
                if (i < alpha.Length) alpha[i] = (byte)c;
            }
            WriteChunk(output, "PLTE", colors);
            if (alpha.Length != 0) WriteChunk(output, "tRNS", alpha);
        }
        var idat = new PngIdatChunkStream(output, StreamingIdatChunkSize);
        WriteColorZlib(idat, width, height, rgba, rgba64, layout, adaptive, token, observer, level, workspace, sampledFilters);
        token.ThrowIfCancellationRequested(); idat.Complete();
        WriteChunk(output, "IEND", Array.Empty<byte>());
    }

    private static void ValidatePngEffort(OfficeImageEncodingEffort effort) {
        if (effort != OfficeImageEncodingEffort.Fast && effort != OfficeImageEncodingEffort.Balanced && effort != OfficeImageEncodingEffort.Smallest)
            throw new ArgumentOutOfRangeException(nameof(effort));
    }

    private static long ColorPayloadBytes(long compressedBytes) =>
        checked(compressedBytes + ((compressedBytes + StreamingIdatChunkSize - 1) / StreamingIdatChunkSize) * 12);

    private static void WriteSrgbChunks(Stream output) {
        WriteChunk(output, "sRGB", new byte[] { 0 });
        var gamma = new byte[4]; WriteBigEndianInt32(gamma, 0, 45455); WriteChunk(output, "gAMA", gamma);
        var chromaticities = new byte[32];
        int[] values = { 31270, 32900, 64000, 33000, 30000, 60000, 15000, 6000 };
        for (int i = 0; i < values.Length; i++) WriteBigEndianInt32(chromaticities, i * 4, values[i]);
        WriteChunk(output, "cHRM", chromaticities);
    }
}
