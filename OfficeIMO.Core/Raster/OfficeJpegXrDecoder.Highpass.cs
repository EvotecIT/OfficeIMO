using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private sealed class HpContext {
        private readonly int _components, _color;
        private readonly CoefficientModel _model;
        private readonly BlockContext _blocks = new();
        private readonly BlockPatterns _patterns;
        private readonly AdaptiveScan _horizontal = new(), _vertical = new(true);

        internal HpContext(int components, int color = 0) {
            _components = components; _color = color; _model = new CoefficientModel(2, components, color); _patterns = new BlockPatterns(components, color);
        }

        internal void Read(Bits bits, int[] output, int offset, int[] patterns, int patternOffset,
                int leftPattern, int topPattern, bool leftEdge, bool topEdge, bool resetScan, bool adapt,
                byte highpassMode, byte[] modelBits, int modelOffset, bool spatialFlex = false, int trim = 0) {
            if (resetScan) { _horizontal.Reset(); _vertical.Reset(); }
            _patterns.Read(bits, patterns, patternOffset, leftPattern, topPattern, leftEdge, topEdge);
            int lumaCount = 0, chromaCount = 0;
            AdaptiveScan scan = highpassMode == 1 ? _vertical : _horizontal;
            for (int c = 0; c < _components; c++) {
                int mask = patterns[patternOffset + c], refinement = _model.Bits[c == 0 ? 0 : 1];
                int blocks = ComponentBlocks(_color, c);
                for (int b = 0; b < blocks; b++) {
                    int block = offset + c * 256 + (blocks == 16 ? HierarchicalScan[b] : b) * 16;
                    if ((mask & (1 << b)) != 0) {
                        int count = _blocks.Read(bits, c != 0), position = 1;
                        if (c == 0) lumaCount += count; else chromaCount += count;
                        for (int i = 0; i < count; i++) {
                            position += _blocks.Runs[i];
                            output[block + scan.Place(position)] = CheckedCoefficient((long)_blocks.Levels[i] << refinement);
                            position++;
                        }
                    }
                    if (spatialFlex) ReadBlockFlex(bits, output, block, refinement, trim);
                }
            }
            modelBits[modelOffset] = (byte)_model.Bits[0]; modelBits[modelOffset + 1] = (byte)_model.Bits[1];
            _model.Update(lumaCount, chromaCount);
            if (adapt) { _blocks.Adapt(); _patterns.Adapt(); }
        }
    }

    internal static BandData ReadFrequencyHp(byte[] bytes, FrameHeader frame, PacketMap packets,
            BandData lp, DclpPlane predicted, DclpPlane? predictedAlpha, CancellationToken cancellation) {
        if (!frame.Frequency || frame.Primary.Bands >= 2) throw new FormatException("JPEG-XR independent HP parsing requires an HP packet.");
        int columns = (frame.Width + frame.Left + frame.Right) / 16, rows = (frame.Height + frame.Top + frame.Bottom) / 16;
        int count = checked(columns * rows), components = frame.Primary.Components;
        if ((long)count * (components + (frame.Alpha ? 1 : 0)) * 256 * 4 + bytes.Length > 256 * 1024 * 1024)
            throw new FormatException("JPEG-XR HP working set exceeds the managed limit.");
        var result = new BandData {
            Primary = new int[checked(count * components * 256)],
            Alpha = frame.Alpha ? new int[checked(count * 256)] : Array.Empty<int>(),
            Quantizers = new int[lp.Quantizers.Length][][],
            AlphaQuantizers = frame.Alpha ? new int[lp.AlphaQuantizers.Length][][] : Array.Empty<int[][]>(),
            QuantizerIndices = new int[count], AlphaQuantizerIndices = frame.Alpha ? new int[count] : Array.Empty<int>(),
            ModelBits = new byte[checked(count * 2)], AlphaModelBits = frame.Alpha ? new byte[checked(count * 2)] : Array.Empty<byte>(),
            Patterns = new int[checked(count * components)], AlphaPatterns = frame.Alpha ? new int[count] : Array.Empty<int>()
        };
        int bands = 4 - frame.Primary.Bands, tile = 0, top = 0;
        for (int tileY = 0; tileY < frame.TileHeights.Length; tileY++) {
            int left = 0, height = frame.TileHeights[tileY];
            for (int tileX = 0; tileX < frame.TileWidths.Length; tileX++, tile++) {
                cancellation.ThrowIfCancellationRequested();
                int width = frame.TileWidths[tileX], packetIndex = tile * bands + 2, start = packets.Offsets[packetIndex];
                var bits = new Bits(bytes, start + 4, packets.Lengths[packetIndex] - 4, cancellation);
                result.Quantizers[tile] = ReadHpQuantization(bits, frame.Primary, lp.Quantizers[tile], out bool reuseLp);
                bool reuseAlphaLp = false;
                if (frame.AlphaPlane != null)
                    result.AlphaQuantizers[tile] = ReadHpQuantization(bits, frame.AlphaPlane, lp.AlphaQuantizers[tile], out reuseAlphaLp);
                var primary = new HpContext(components, frame.Primary.Color); HpContext? alpha = frame.Alpha ? new HpContext(1) : null;
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    int mb = (top + y) * columns + left + x; bool adapt = x == width - 1 || x % 16 == 0;
                    result.QuantizerIndices[mb] = reuseLp ? lp.QuantizerIndices[mb] : ReadQuantizerIndex(bits, result.Quantizers[tile].Length);
                    primary.Read(bits, result.Primary, mb * components * 256, result.Patterns, mb * components,
                        (mb - 1) * components, (mb - columns) * components, x == 0, y == 0, x % 16 == 0, adapt,
                        predicted.HighpassModes[mb], result.ModelBits, mb * 2);
                    if (alpha != null && frame.AlphaPlane!.Bands < 2) {
                        result.AlphaQuantizerIndices[mb] = reuseAlphaLp ? lp.AlphaQuantizerIndices[mb] : ReadQuantizerIndex(bits, result.AlphaQuantizers[tile].Length);
                        alpha.Read(bits, result.Alpha, mb * 256, result.AlphaPatterns, mb, mb - 1, mb - columns,
                            x == 0, y == 0, x % 16 == 0, adapt, predictedAlpha!.HighpassModes[mb], result.AlphaModelBits, mb * 2);
                    }
                }
                bits.AlignZero(); left += width;
            }
            top += height;
        }
        return result;
    }

    private static int[][] ReadHpQuantization(Bits bits, PlaneHeader plane, int[][] lpQuantizers, out bool reuseLp) {
        reuseLp = false;
        if (plane.Bands >= 2) return Array.Empty<int[]>();
        if (plane.HpQuant != null) return new[] { plane.HpQuant };
        if (bits.Flag()) { reuseLp = true; return lpQuantizers; }
        int sets = (int)bits.Read(4) + 1;
        var quantizers = new int[sets][];
        for (int i = 0; i < sets; i++) quantizers[i] = ReadQuantization(bits, plane.Components);
        return quantizers;
    }
}
