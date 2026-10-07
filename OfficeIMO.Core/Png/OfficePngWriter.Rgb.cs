using System;
using System.Threading;
#if NET8_0_OR_GREATER
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
#endif

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    // Select a lossless sample layout before writing any bytes. A transparent
    // pixel anywhere in the image requires preserving every RGBA channel.
    private static int SelectOptimalColorType(byte[] rgba, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        bool bilevel = true;
        for (int offset = 0; offset < rgba.Length; offset += 4) {
            if ((offset & 4095) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngColorSelectionBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            if (rgba[offset + 3] != 255) return 6;
            if (bilevel) {
                byte gray = rgba[offset];
                if ((gray != 0 && gray != 255) || rgba[offset + 1] != gray || rgba[offset + 2] != gray)
                    bilevel = false;
            }
        }
        return bilevel ? 0 : 2;
    }

    // The existing RGBA filters already predict the same R/G/B channels as RGB
    // filters. Constant opaque alpha contributes zero to Up/Paeth scores. Compact
    // the selected row in place; source pixels and scratch allocation stay intact.
    private static void CompactOpaqueRgbRow(byte[] row, int rgbaStride,
        CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        int target = 1;
        int source = 1;
#if NET8_0_OR_GREATER
        if (Ssse3.IsSupported) {
            var channels = Vector128.Create((byte)0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 13, 14,
                byte.MaxValue, byte.MaxValue, byte.MaxValue, byte.MaxValue);
            for (; source <= rgbaStride - 15; source += 16, target += 12) {
                if (((source - 1) & 4095) == 0) {
                    checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
                    cancellationToken.ThrowIfCancellationRequested();
                }
                // Load before the overlapping store. The four unused store bytes
                // remain inside the RGBA scratch and are replaced by the next block
                // or scalar tail; only the compacted RGB prefix is emitted.
                var pixels = Vector128.LoadUnsafe(ref row[source]);
                Ssse3.Shuffle(pixels, channels).StoreUnsafe(ref row[target]);
            }
        }
#endif
        for (; source <= rgbaStride; source += 4) {
            if (((source - 1) & 4095) == 0) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngFilteringBlock);
                cancellationToken.ThrowIfCancellationRequested();
            }
            row[target++] = row[source];
            row[target++] = row[source + 1];
            row[target++] = row[source + 2];
        }
    }
}
