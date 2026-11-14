using System.IO.Compression;

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
        PngColorWorkspace? workspace = null, bool sampledFilters = false) {
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
                    for (int i = 0, step = sampledFilters ? Math.Max(1, stride / 128) : 1; i < stride; i += step) {
                        if (sampledFilters) token.ThrowIfCancellationRequested();
                        else CheckColorBlock(i, token, observer);
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
                        for (int blockStart = 0; blockStart < stride; blockStart += 4096) {
                            CheckColorBlock(blockStart, token, observer);
                            ApplyColorFilter(filter, current, previous, selected, bpp, blockStart, Math.Min(stride, blockStart + 4096));
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

    // Pixels are copied in blocks of 1024 so the filtering checkpoint still follows every 4096 source bytes.
    private static void FillColorRow(byte[] row, int width, int y, byte[]? rgba, ushort[]? rgba64,
        PngColorLayout layout, CancellationToken token, Action<OfficeRasterEncodingCheckpoint>? observer) {
        int rowSource = checked(y * width * 4);
        if (layout.Depth < 8) Array.Clear(row, 0, layout.RowBytes(width));
        bool packed = layout.Type == 3 || layout.Depth < 8;
        uint lastKey = 0;
        int lastIndex = -1;
        for (int blockStart = 0; blockStart < width; blockStart += 1024) {
            CheckColorBlock(blockStart * 4, token, observer);
            int blockEnd = Math.Min(width, blockStart + 1024);
            int source = rowSource + blockStart * 4;
            if (packed) {
                for (int x = blockStart; x < blockEnd; x++, source += 4) {
                    int value;
                    if (layout.Type == 3) {
                        uint key = RgbaKey(rgba!, source);
                        if (lastIndex < 0 || key != lastKey) { lastIndex = layout.Palette![key]; lastKey = key; }
                        value = lastIndex;
                    } else value = rgba![source] * ((1 << layout.Depth) - 1) / 255;
                    int bit = x * layout.Depth;
                    if (layout.Depth == 8) row[x] = (byte)value;
                    else row[bit >> 3] |= (byte)(value << (8 - layout.Depth - (bit & 7)));
                }
            } else if (rgba64 != null) {
                int target = blockStart * layout.Channels * 2;
                for (int x = blockStart; x < blockEnd; x++, source += 4) {
                    for (int c = 0; c < layout.Channels; c++) {
                        ushort value = rgba64[source + (layout.Type == 4 && c == 1 ? 3 : c)];
                        row[target++] = (byte)(value >> 8); row[target++] = (byte)value;
                    }
                }
            } else if (layout.Type == 6) {
                Buffer.BlockCopy(rgba!, source, row, blockStart * 4, (blockEnd - blockStart) * 4);
            } else if (layout.Type == 2) {
                for (int x = blockStart, target = blockStart * 3; x < blockEnd; x++, source += 4, target += 3) {
                    row[target] = rgba![source]; row[target + 1] = rgba[source + 1]; row[target + 2] = rgba[source + 2];
                }
            } else if (layout.Type == 4) {
                for (int x = blockStart, target = blockStart * 2; x < blockEnd; x++, source += 4, target += 2) {
                    row[target] = rgba![source]; row[target + 1] = rgba[source + 3];
                }
            } else {
                for (int x = blockStart; x < blockEnd; x++, source += 4) row[x] = rgba![source];
            }
        }
    }

    // Writes filtered bytes [start, end) after the filter-type byte, with the same predictions as the
    // PNG reference filters. Sub, Up and Average have no carried state and run on vectors.
    private static void ApplyColorFilter(int filter, byte[] current, byte[] previous, byte[] selected, int bpp, int start, int end) {
        int i = start;
        for (; i < end && i < bpp; i++) {
            int up = previous[i];
            int prediction = filter == 1 ? 0 : filter == 2 ? up : filter == 3 ? up / 2 : ColorPaeth(0, up, 0);
            selected[i + 1] = unchecked((byte)(current[i] - prediction));
        }
#if NET8_0_OR_GREATER
        if (filter != 4 && System.Numerics.Vector.IsHardwareAccelerated) {
            int width = System.Numerics.Vector<byte>.Count;
            for (; i <= end - width; i += width) {
                var value = new System.Numerics.Vector<byte>(current, i);
                var up = new System.Numerics.Vector<byte>(previous, i);
                System.Numerics.Vector<byte> prediction;
                if (filter == 2) prediction = up;
                else {
                    var left = new System.Numerics.Vector<byte>(current, i - bpp);
                    // floor((left + up) / 2) without leaving eight bits.
                    prediction = filter == 1 ? left : (left & up) + System.Numerics.Vector.ShiftRightLogical(left ^ up, 1);
                }
                (value - prediction).CopyTo(selected, i + 1);
            }
        }
#endif
        switch (filter) {
            case 1:
                for (; i < end; i++) selected[i + 1] = unchecked((byte)(current[i] - current[i - bpp]));
                break;
            case 2:
                for (; i < end; i++) selected[i + 1] = unchecked((byte)(current[i] - previous[i]));
                break;
            case 3:
                for (; i < end; i++) selected[i + 1] = unchecked((byte)(current[i] - ((current[i - bpp] + previous[i]) >> 1)));
                break;
            default:
                for (; i < end; i++) selected[i + 1] = unchecked((byte)(current[i] - ColorPaeth(current[i - bpp], previous[i], previous[i - bpp])));
                break;
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
