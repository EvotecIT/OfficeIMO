using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private sealed class AdaptiveScan {
        private readonly int[] _order, _totals = new int[16];
        internal AdaptiveScan(bool vertical = false) { _order = (int[])(vertical ? VerticalScan : HorizontalScan).Clone(); Reset(); }
        internal void Reset() { for (int i = 1; i < 16; i++) _totals[i] = 34 - 2 * i; }
        internal int Place(int index) {
            if (index < 1 || index > 15) throw new FormatException("JPEG-XR scan position exceeds its block.");
            int output = _order[index]; _totals[index]++;
            if (index > 1 && _totals[index] > _totals[index - 1]) {
                int temp = _totals[index]; _totals[index] = _totals[index - 1]; _totals[index - 1] = temp;
                temp = _order[index]; _order[index] = _order[index - 1]; _order[index - 1] = temp;
            }
            return output;
        }
    }

    private sealed partial class LpContext {
        private readonly int _components, _color;
        private readonly CoefficientModel _model;
        private readonly BlockContext _blocks = new();
        private readonly AdaptiveScan _scan = new();
        private int _zeroCount = 1, _maximumCount = 1;
        internal LpContext(int components, int color = 0) { _components = components; _color = color; _model = new CoefficientModel(1, components, color); }

        internal void Read(Bits bits, int[] output, int offset, bool resetScan, bool adapt) {
            if (resetScan) _scan.Reset();
            bool reduced = _color == 1 || _color == 2;
            int maximum = reduced ? 3 : 7;
            int presence;
            if (_components == 1) presence = (int)bits.Read(1);
            else {
                presence = _zeroCount <= 0 || _maximumCount < 0 ? (reduced ? SubsampledLowpassPresence : LowpassPresence).Read(bits) : (int)bits.Read(reduced ? 2 : 3);
                if ((_zeroCount <= 0 || _maximumCount < 0) && _maximumCount < _zeroCount) presence = maximum - presence;
                _zeroCount = Math.Max(-8, Math.Min(7, _zeroCount + 1 - (presence == 0 ? 4 : 0)));
                _maximumCount = Math.Max(-8, Math.Min(7, _maximumCount + 1 - (presence == maximum ? 4 : 0)));
            }
            int lumaCount = 0, chromaCount = 0;
            for (int channel = 0; channel < (reduced ? 2 : _components); channel++) {
                if (reduced && channel == 1) {
                    chromaCount = ReadReducedChroma(bits, output, offset, (presence & 2) != 0);
                    continue;
                }
                int start = offset + channel * 16, nonzero = 0;
                if ((presence & (1 << channel)) != 0) {
                    nonzero = _blocks.Read(bits, channel != 0);
                    int position = 1;
                    for (int i = 0; i < nonzero; i++) {
                        position += _blocks.Runs[i];
                        output[start + _scan.Place(position)] = _blocks.Levels[i]; position++;
                    }
                }
                if (channel == 0) lumaCount += nonzero; else chromaCount += nonzero;
                int refinement = _model.Bits[channel == 0 ? 0 : 1];
                if (refinement != 0) for (int k = 1; k < 16; k++) {
                    int index = start + TransposedScan[k], original = output[index];
                    uint tail = bits.Read(refinement);
                    long value = original == 0 ? tail : (Math.Abs((long)original) << refinement) + tail;
                    if (value > int.MaxValue) throw new FormatException("JPEG-XR LP coefficient exceeds the managed range.");
                    bool negative = original < 0 || original == 0 && value != 0 && bits.Flag();
                    output[index] = negative ? -(int)value : (int)value;
                }
            }
            _model.Update(lumaCount, chromaCount);
            if (adapt) _blocks.Adapt();
        }
    }

    private static readonly Codebook SubsampledLowpassPresence = new("0", "10", "110", "111");

    internal static BandData ReadFrequencyLp(byte[] bytes, FrameHeader frame, PacketMap packets, BandData dc,
            CancellationToken cancellation) {
        if (!frame.Frequency || frame.Primary.Bands == 3) throw new FormatException("JPEG-XR independent LP parsing requires an LP packet.");
        int columns = (frame.Width + frame.Left + frame.Right) / 16, rows = (frame.Height + frame.Top + frame.Bottom) / 16;
        int count = checked(columns * rows), components = frame.Primary.Components;
        if ((long)count * (components + (frame.Alpha ? 1 : 0)) * 16 * 4 + bytes.Length > 256 * 1024 * 1024)
            throw new FormatException("JPEG-XR LP working set exceeds the managed limit.");
        var result = new BandData {
            Primary = new int[checked(count * components * 16)],
            Alpha = frame.Alpha ? new int[checked(count * 16)] : Array.Empty<int>(),
            Quantizers = new int[dc.Quantizers.Length][][],
            AlphaQuantizers = frame.Alpha ? new int[dc.AlphaQuantizers.Length][][] : Array.Empty<int[][]>(),
            QuantizerIndices = new int[count], AlphaQuantizerIndices = frame.Alpha ? new int[count] : Array.Empty<int>()
        };
        int bands = 4 - frame.Primary.Bands, tile = 0, top = 0;
        for (int tileY = 0; tileY < frame.TileHeights.Length; tileY++) {
            int left = 0, height = frame.TileHeights[tileY];
            for (int tileX = 0; tileX < frame.TileWidths.Length; tileX++, tile++) {
                cancellation.ThrowIfCancellationRequested();
                int width = frame.TileWidths[tileX], packetIndex = tile * bands + 1, start = packets.Offsets[packetIndex];
                var bits = new Bits(bytes, start + 4, packets.Lengths[packetIndex] - 4, cancellation);
                result.Quantizers[tile] = ReadLpQuantization(bits, frame.Primary, dc.Quantizers[tile][0]);
                if (frame.AlphaPlane != null)
                    result.AlphaQuantizers[tile] = ReadLpQuantization(bits, frame.AlphaPlane, dc.AlphaQuantizers[tile][0]);
                var primary = new LpContext(components, frame.Primary.Color); LpContext? alpha = frame.Alpha ? new LpContext(1) : null;
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    int macroblock = (top + y) * columns + left + x; bool adapt = x == width - 1 || x % 16 == 0;
                    result.QuantizerIndices[macroblock] = ReadQuantizerIndex(bits, result.Quantizers[tile].Length);
                    primary.Read(bits, result.Primary, macroblock * components * 16, x % 16 == 0, adapt);
                    if (alpha != null && frame.AlphaPlane!.Bands != 3) {
                        result.AlphaQuantizerIndices[macroblock] = ReadQuantizerIndex(bits, result.AlphaQuantizers[tile].Length);
                        alpha.Read(bits, result.Alpha, macroblock * 16, x % 16 == 0, adapt);
                    }
                }
                bits.AlignZero(); left += width;
            }
            top += height;
        }
        return result;
    }

    private static int[][] ReadLpQuantization(Bits bits, PlaneHeader plane, int[] dcQuantization) {
        if (plane.Bands == 3) return Array.Empty<int[]>();
        if (plane.LpQuant != null) return new[] { plane.LpQuant };
        if (bits.Flag()) return new[] { dcQuantization };
        int sets = (int)bits.Read(4) + 1;
        var quantizers = new int[sets][];
        for (int i = 0; i < sets; i++) quantizers[i] = ReadQuantization(bits, plane.Components);
        return quantizers;
    }

    private static int ReadQuantizerIndex(Bits bits, int count) {
        if (count <= 1) return 0;
        if (!bits.Flag()) return 0;
        int width = 0;
        for (int remaining = count - 2; remaining > 0; remaining >>= 1) width++;
        int index = (int)bits.Read(width) + 1;
        if (index >= count) throw new FormatException("JPEG-XR quantization index exceeds its table.");
        return index;
    }
}
