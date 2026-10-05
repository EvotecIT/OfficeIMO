using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {

    // The frequency DC packet can be decoded independently. Spatial mode needs
    // the LP/HP contexts between successive DC macroblocks.
    internal static BandData ReadFrequencyDc(byte[] bytes, FrameHeader frame, PacketMap packets,
            CancellationToken cancellation) {
        if (!frame.Frequency) throw new FormatException("JPEG-XR independent DC parsing requires frequency packets.");
        int columns = (frame.Width + frame.Left + frame.Right) / 16;
        int rows = (frame.Height + frame.Top + frame.Bottom) / 16;
        int count = checked(columns * rows), components = frame.Primary.Components;
        int tileCount = checked(frame.TileWidths.Length * frame.TileHeights.Length);
        var result = new BandData {
            Primary = new int[checked(count * components)],
            Alpha = frame.Alpha ? new int[count] : Array.Empty<int>(),
            Quantizers = new int[tileCount][][], AlphaQuantizers = frame.Alpha ? new int[tileCount][][] : Array.Empty<int[][]>()
        };
        int bands = 4 - frame.Primary.Bands, tile = 0, top = 0;
        for (int tileY = 0; tileY < frame.TileHeights.Length; tileY++) {
            int left = 0, height = frame.TileHeights[tileY];
            for (int tileX = 0; tileX < frame.TileWidths.Length; tileX++, tile++) {
                cancellation.ThrowIfCancellationRequested();
                int width = frame.TileWidths[tileX], packetIndex = tile * bands;
                int start = packets.Offsets[packetIndex];
                var bits = new Bits(bytes, start + 4, packets.Lengths[packetIndex] - 4, cancellation);
                result.Quantizers[tile] = new[] { frame.Primary.DcQuant ?? ReadQuantization(bits, components) };
                if (frame.AlphaPlane != null)
                    result.AlphaQuantizers[tile] = new[] { frame.AlphaPlane.DcQuant ?? ReadQuantization(bits, 1) };
                var primary = new DcContext(components);
                DcContext? alpha = frame.Alpha ? new DcContext(1) : null;
                for (int y = 0; y < height; y++) {
                    for (int x = 0; x < width; x++) {
                        int macroblock = (top + y) * columns + left + x;
                        bool adapt = x == width - 1 || x % 16 == 0;
                        primary.Read(bits, result.Primary, macroblock * components, adapt);
                        if (alpha != null) {
                            alpha.Read(bits, result.Alpha, macroblock, adapt);
                        }
                    }
                }
                bits.AlignZero(); left += width;
            }
            top += height;
        }
        return result;
    }
}
