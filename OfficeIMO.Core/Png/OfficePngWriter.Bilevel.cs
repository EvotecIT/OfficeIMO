using System;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    // Scan cleanup produces RGBA pixels too. Pack only exact opaque black/white;
    // intermediate gray, colored and transparent pixels must keep all channels.
    private static bool IsOpaqueBilevel(byte[] rgba, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        for (int offset = 0; offset < rgba.Length; offset += 4) {
            if ((offset & 4095) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngColorSelectionBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            byte gray = rgba[offset];
            if ((gray != 0 && gray != 255) || rgba[offset + 1] != gray
                || rgba[offset + 2] != gray || rgba[offset + 3] != 255) return false;
        }
        return true;
    }

    private static void FilterBilevelRow(byte[] rgba, int rgbaOffset, PngFilteringWorkspace workspace,
        int y, bool adaptiveFiltering, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        byte[] rows = workspace.BilevelRows!;
        int stride = workspace.Stride;
        int currentOffset = (y & 1) * stride;
        int previousOffset = ((y & 1) ^ 1) * stride;
        int width = workspace.RgbaStride / 4;
        for (int x = 0; x < width;) {
            if ((x & 1023) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            int byteIndex = x >> 3;
            int end = x + Math.Min(8, width - x);
            int value = 0;
            for (int bit = 7; x < end; x++, bit--)
                value |= (rgba[rgbaOffset + x * 4] >> 7) << bit;
            rows[currentOffset + byteIndex] = (byte)value;
        }

        byte[] filtered = workspace.Row;
        if (!adaptiveFiltering) {
            filtered[0] = 0;
            Buffer.BlockCopy(rows, currentOffset, filtered, 1, stride);
        } else if (y == 0) {
            filtered[0] = 1;
            filtered[1] = rows[currentOffset];
            for (int index = 1; index < stride; index++) {
                CheckBilevelFilterCancellation(index, cancellationToken, checkpointObserver);
                filtered[index + 1] = unchecked((byte)(rows[currentOffset + index] - rows[currentOffset + index - 1]));
            }
        } else {
            long upScore = FilterUp(rows, currentOffset, previousOffset, stride,
                filtered, 1, cancellationToken, checkpointObserver);
            long paethScore = 0L;
            for (int index = 0; index < stride; index++) {
                CheckBilevelFilterCancellation(index, cancellationToken, checkpointObserver);
                int left = index > 0 ? rows[currentOffset + index - 1] : 0;
                int above = rows[previousOffset + index];
                int upperLeft = index > 0 ? rows[previousOffset + index - 1] : 0;
                byte value = unchecked((byte)(rows[currentOffset + index] - PaethPredictor(left, above, upperLeft)));
                workspace.Paeth[index] = value;
                paethScore += Math.Abs((int)(sbyte)value);
            }
            if (paethScore < upScore) {
                filtered[0] = 4;
                Buffer.BlockCopy(workspace.Paeth, 0, filtered, 1, stride);
            } else filtered[0] = 2;
        }
    }

    private static void CheckBilevelFilterCancellation(int index, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        if ((index & 4095) != 0) return;
        checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
        cancellationToken.ThrowIfCancellationRequested();
    }
}
