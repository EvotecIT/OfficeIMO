using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    private sealed class PngColorWorkspace {
        internal PngColorWorkspace(int maximumRowBytes) {
            Current = new byte[maximumRowBytes]; Previous = new byte[maximumRowBytes];
            Selected = new byte[checked(maximumRowBytes + 1)];
        }
        internal byte[] Current, Previous, Selected;
    }

    private static void WriteColorZlib(Stream destination, int width, int height, byte[]? rgba,
        ushort[]? rgba64, PngColorLayout layout, bool adaptive, CancellationToken token,
        Action<OfficeRasterEncodingCheckpoint>? observer, CompressionLevel? level = null,
        PngColorWorkspace? workspace = null) {
        int stride = layout.RowBytes(width), bpp = layout.FilterBytesPerPixel;
        EnsureColorWorkingSet(rgba?.LongLength ?? checked(rgba64!.LongLength * 2L), stride);
        workspace ??= new PngColorWorkspace(stride);
        byte[] current = workspace.Current, previous = workspace.Previous;
        byte[] selected = workspace.Selected;
        Array.Clear(previous, 0, stride);
        destination.WriteByte(0x78); destination.WriteByte(0x9C);
        uint adlerA = 1, adlerB = 0;
#if NET8_0_OR_GREATER
        CompressionLevel optimal = CompressionLevel.SmallestSize;
#else
        CompressionLevel optimal = CompressionLevel.Optimal;
#endif
        using (var deflate = new DeflateStream(destination, level ?? optimal, leaveOpen: true)) {
            for (int y = 0; y < height; y++) {
                observer?.Invoke(OfficeRasterEncodingCheckpoint.PngCompressionRow); token.ThrowIfCancellationRequested();
                FillColorRow(current, width, y, rgba, rgba64, layout, token, observer);
                selected[0] = 0; Buffer.BlockCopy(current, 0, selected, 1, stride);
                if (adaptive) {
                    long none = 0, sub = 0, upScore = 0, average = 0, paeth = 0;
                    // Score all predictors in one scan. Replaying just the winning
                    // filter avoids four candidate rows and repeated source walks.
                    for (int i = 0; i < stride; i++) {
                        CheckColorBlock(i, token, observer);
                        int value = current[i], left = i >= bpp ? current[i - bpp] : 0;
                        int up = previous[i], upperLeft = i >= bpp ? previous[i - bpp] : 0;
                        none += SignedMagnitude(value);
                        sub += SignedMagnitude(value - left);
                        upScore += SignedMagnitude(value - up);
                        average += SignedMagnitude(value - (left + up) / 2);
                        paeth += SignedMagnitude(value - ColorPaeth(left, up, upperLeft));
                    }
                    long best = none; int filter = 0;
                    if (sub < best) { best = sub; filter = 1; }
                    if (upScore < best) { best = upScore; filter = 2; }
                    if (average < best) { best = average; filter = 3; }
                    if (paeth < best) filter = 4;
                    if (filter != 0) {
                        selected[0] = (byte)filter;
                        for (int i = 0; i < stride; i++) {
                            CheckColorBlock(i, token, observer);
                            int left = i >= bpp ? current[i - bpp] : 0;
                            int up = previous[i], upperLeft = i >= bpp ? previous[i - bpp] : 0;
                            int prediction = filter == 1 ? left : filter == 2 ? up : filter == 3 ? (left + up) / 2 : ColorPaeth(left, up, upperLeft);
                            selected[i + 1] = unchecked((byte)(current[i] - prediction));
                        }
                    }
                }
                deflate.Write(selected, 0, stride + 1);
                UpdateAdler32(selected, 0, stride + 1, ref adlerA, ref adlerB, token, observer);
                var rowSwap = previous; previous = current; current = rowSwap;
            }
        }
        WriteAdler32(destination, (adlerB << 16) | adlerA);
    }

    private static void FillColorRow(byte[] row, int width, int y, byte[]? rgba, ushort[]? rgba64,
        PngColorLayout layout, CancellationToken token, Action<OfficeRasterEncodingCheckpoint>? observer) {
        int source = checked(y * width * 4), target = 0;
        if (layout.Depth < 8) Array.Clear(row, 0, layout.RowBytes(width));
        for (int x = 0; x < width; x++, source += 4) {
            CheckColorBlock(x * 4, token, observer);
            if (layout.Type == 3 || layout.Depth < 8) {
                int value = layout.Type == 3 ? layout.Palette![RgbaKey(rgba!, source)] : rgba![source] * ((1 << layout.Depth) - 1) / 255;
                int bit = x * layout.Depth;
                if (layout.Depth == 8) row[x] = (byte)value;
                else row[bit >> 3] |= (byte)(value << (8 - layout.Depth - (bit & 7)));
            } else {
                for (int c = 0; c < layout.Channels; c++) {
                    int channel = layout.Type == 4 && c == 1 ? 3 : c;
                    if (rgba64 != null) {
                        ushort value = rgba64[source + channel]; row[target++] = (byte)(value >> 8); row[target++] = (byte)value;
                    } else row[target++] = rgba![source + channel];
                }
            }
        }
    }

    private static void CheckColorBlock(int index, CancellationToken token, Action<OfficeRasterEncodingCheckpoint>? observer) {
        if ((index & 4095) == 0) { observer?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock); token.ThrowIfCancellationRequested(); }
    }

    private static int SignedMagnitude(int value) => Math.Abs((int)unchecked((sbyte)value));

    private static void EnsureColorWorkingSet(long sourceBytes, int rowBytes) {
        if (sourceBytes + (long)rowBytes * 3 + 65536 > OfficeRasterGuards.MaximumDecodedBytes)
            throw new ArgumentException("PNG row filtering exceeds the managed working-set limit.");
    }

    private static int ColorPaeth(int left, int above, int upperLeft) {
        int p = left + above - upperLeft;
        int a = Math.Abs(p - left), b = Math.Abs(p - above), c = Math.Abs(p - upperLeft);
        return a <= b && a <= c ? left : b <= c ? above : upperLeft;
    }
}
