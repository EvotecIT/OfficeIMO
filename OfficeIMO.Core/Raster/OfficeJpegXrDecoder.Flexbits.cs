using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    // Refines the owned HP arrays in place; packet contexts and decoded DC/LP
    // values are independent of this last frequency pass.
    internal static BandData ReadFrequencyFlex(byte[] bytes, FrameHeader frame, PacketMap packets,
            BandData highpass, CancellationToken cancellation) {
        if (!frame.Frequency || frame.Primary.Bands != 0)
            throw new FormatException("JPEG-XR flexbits parsing requires frequency all-band packets.");
        int columns = (frame.Width + frame.Left + frame.Right) / 16;
        var result = highpass;
        int tile = 0, top = 0;
        for (int tileY = 0; tileY < frame.TileHeights.Length; tileY++) {
            int left = 0, height = frame.TileHeights[tileY];
            for (int tileX = 0; tileX < frame.TileWidths.Length; tileX++, tile++) {
                cancellation.ThrowIfCancellationRequested();
                int width = frame.TileWidths[tileX], packetIndex = tile * 4 + 3, start = packets.Offsets[packetIndex];
                var bits = new Bits(bytes, start + 4, packets.Lengths[packetIndex] - 4, cancellation);
                int trim = frame.Trim ? (int)bits.Read(4) : 0;
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    int mb = (top + y) * columns + left + x;
                    ReadMacroblockFlex(bits, result.Primary, mb * frame.Primary.Components * 256, frame.Primary.Components,
                        highpass.ModelBits, mb * 2, trim, frame.Primary.Color);
                    if (frame.AlphaPlane != null && frame.AlphaPlane.Bands == 0) {
                        ReadMacroblockFlex(bits, result.Alpha, mb * 256, 1, highpass.AlphaModelBits, mb * 2, trim);
                    }
                }
                bits.AlignZero(); left += width;
            }
            top += height;
        }
        return result;
    }

    private static void ReadMacroblockFlex(Bits bits, int[] coefficients, int offset, int components,
            byte[] modelBits, int modelOffset, int trim, int color = 0) {
        for (int c = 0; c < components; c++) {
            int model = modelBits[modelOffset + (c == 0 ? 0 : 1)];
            int blocks = ComponentBlocks(color, c);
            for (int b = 0; b < blocks; b++) ReadBlockFlex(bits, coefficients, offset + c * 256 + (blocks == 16 ? HierarchicalScan[b] : b) * 16, model, trim);
        }
    }

    private static int ComponentBlocks(int color, int component) => component == 0 ? 16
        : color == 1 ? 4 : color == 2 ? 8 : 16;

    private static void ReadBlockFlex(Bits bits, int[] coefficients, int offset, int model, int trim) {
        int remaining = Math.Max(0, model - trim);
        if (remaining == 0) return;
        for (int k = 1; k < 16; k++) {
            int index = offset + TransposedScan[k], original = coefficients[index], tail = (int)bits.Read(remaining) << trim;
            bool negative = original < 0 || original == 0 && tail != 0 && bits.Flag();
            coefficients[index] = CheckedCoefficient((long)original + (negative ? -tail : tail));
        }
    }

    internal static void ReconstructHighpass(FrameHeader frame, BandData highpass, DclpPlane predicted,
            DclpPlane? predictedAlpha, CancellationToken cancellation) {
        int columns = (frame.Width + frame.Left + frame.Right) / 16, tile = 0, top = 0;
        for (int tileY = 0; tileY < frame.TileHeights.Length; tileY++) {
            int left = 0, height = frame.TileHeights[tileY];
            for (int tileX = 0; tileX < frame.TileWidths.Length; tileX++, tile++) {
                int width = frame.TileWidths[tileX];
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    cancellation.ThrowIfCancellationRequested();
                    int mb = (top + y) * columns + left + x;
                    ReconstructMacroblockHp(highpass.Primary, mb * frame.Primary.Components * 256, frame.Primary,
                        highpass.Quantizers[tile][highpass.QuantizerIndices[mb]], predicted.HighpassModes[mb]);
                    if (frame.AlphaPlane != null && frame.AlphaPlane.Bands < 2)
                        ReconstructMacroblockHp(highpass.Alpha, mb * 256, frame.AlphaPlane,
                            highpass.AlphaQuantizers[tile][highpass.AlphaQuantizerIndices[mb]], predictedAlpha!.HighpassModes[mb]);
                }
                left += width;
            }
            top += height;
        }
    }

    private static void ReconstructMacroblockHp(int[] coefficients, int offset, PlaneHeader plane, int[] quantizers, byte mode) {
        for (int c = 0; c < plane.Components; c++) {
            int start = offset + c * 256, scale = QuantMap(quantizers[c], 1, plane.Scaled);
            int blocks = ComponentBlocks(plane.Color, c), width = blocks == 16 ? 4 : 2;
            for (int b = 0; b < blocks; b++) for (int k = 1; k < 16; k++)
                coefficients[start + b * 16 + k] = CheckedCoefficient((long)coefficients[start + b * 16 + k] * scale);
            if (mode == 2) continue;
            for (int b = mode == 0 ? 1 : width; b < blocks; b++) {
                if (mode == 0 && b % width == 0) continue;
                int neighbour = mode == 0 ? b - 1 : b - width;
                for (int k = mode == 0 ? 4 : 1; k <= (mode == 0 ? 12 : 3); k += mode == 0 ? 4 : 1) {
                    int index = start + b * 16 + k;
                    coefficients[index] = CheckedCoefficient((long)coefficients[index] + coefficients[start + neighbour * 16 + k]);
                }
            }
        }
    }
}
