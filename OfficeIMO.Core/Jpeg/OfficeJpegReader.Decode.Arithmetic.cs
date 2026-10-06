using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private sealed class ArithmeticConditioning {
        public readonly byte[] Lower = new byte[16];
        public readonly byte[] Upper = new byte[16];
        public readonly byte[] AcBoundary = new byte[16];
        public ArithmeticConditioning() {
            for (int i = 0; i < 16; i++) { Upper[i] = 1; AcBoundary[i] = 5; }
        }
        public void Read(OfficeByteView bytes) {
            if ((bytes.Length & 1) != 0) throw new FormatException("Invalid JPEG arithmetic conditioning segment.");
            for (int at = 0; at < bytes.Length; at += 2) {
                int selector = bytes[at], value = bytes[at + 1], id = selector & 15;
                if (selector < 16) {
                    int lower = value & 15, upper = value >> 4;
                    if (lower > upper) throw new FormatException("Invalid JPEG DC conditioning bounds.");
                    Lower[id] = (byte)lower; Upper[id] = (byte)upper;
                } else if (selector < 32 && value <= 63) AcBoundary[id] = (byte)value;
                else throw new FormatException("Invalid JPEG arithmetic conditioning selector.");
            }
        }
    }

    private static void DecodeArithmeticSequential(OfficeByteView data, ScanHeader scan, JpegFrame frame,
        BaselineState state, int[][] quantization, ArithmeticConditioning conditioning,
        int restartInterval, CancellationToken token) {
        if (scan.Ss != 0 || scan.Se != 63 || scan.Ah != 0 || scan.Al != 0)
            throw new FormatException("Invalid sequential arithmetic JPEG scan.");
        // Sixteen shared conditioning destinations, as distinct from the four
        // quantization/Huffman destinations. Statistics reset at every scan/RST.
        byte[][] dc = new byte[16][], ac = new byte[16][];
        int[] context = new int[frame.ComponentCount];
        foreach (int ci in scan.ComponentIndices) {
            var component = frame.Components[ci];
            if (quantization[component.QuantId] is null) throw new FormatException("Missing JPEG quantization table.");
            if (state.DecodedComponents[ci]) throw new FormatException("Repeated sequential JPEG component scan.");
            dc[component.DcTable] ??= new byte[49]; ac[component.AcTable] ??= new byte[245];
            state.Components[ci].PrevDc = 0;
        }
        var reader = new ArithmeticReader(data, token);
        bool single = scan.ComponentIndices.Length == 1;
        var first = frame.Components[scan.ComponentIndices[0]];
        int columns = single ? GetNonInterleavedBlockCount(frame.Width, first.H, frame.MaxH) : state.McuCols;
        int rows = single ? GetNonInterleavedBlockCount(frame.Height, first.V, frame.MaxV) : state.McuRows;
        int mcu = 0;
        for (int y = 0; y < rows; y++) for (int x = 0; x < columns; x++, mcu++) {
            token.ThrowIfCancellationRequested();
            if (restartInterval > 0 && mcu > 0 && mcu % restartInterval == 0) {
                reader.Restart();
                for (int t = 0; t < 16; t++) {
                    if (dc[t] != null) Array.Clear(dc[t], 0, dc[t].Length);
                    if (ac[t] != null) Array.Clear(ac[t], 0, ac[t].Length);
                }
                Array.Clear(context, 0, context.Length);
                foreach (int ci in scan.ComponentIndices) state.Components[ci].PrevDc = 0;
            }
            foreach (int ci in scan.ComponentIndices) {
                var component = frame.Components[ci]; var pixels = state.Components[ci];
                int blocks = single ? 1 : component.H * component.V;
                for (int b = 0; b < blocks; b++) {
                    int[] coefficients = pixels.BlockCoeffs; Array.Clear(coefficients, 0, coefficients.Length);
                    byte[] dcBins = dc[component.DcTable], acBins = ac[component.AcTable];
                    DecodeArithmeticDc(reader, dcBins, conditioning, component.DcTable,
                        ref context[ci], ref pixels.PrevDc);
                    coefficients[0] = checked(pixels.PrevDc * quantization[component.QuantId][0]);
                    for (int k = 1; k < 64; k++) {
                        int at = 3 * (k - 1);
                        if (reader.Decode(ref acBins[at]) != 0) break;
                        while (reader.Decode(ref acBins[at + 1]) == 0) {
                            if (++k == 64) throw new FormatException("Arithmetic JPEG coefficient run exceeds the block.");
                            at += 3;
                        }
                        int sign = reader.DecodeFixed();
                        int magnitude = DecodeArithmeticMagnitude(reader, acBins, at + 2, at + 2,
                            k <= conditioning.AcBoundary[component.AcTable] ? 189 : 217);
                        int natural = ZigZag[k];
                        coefficients[natural] = checked((sign == 0 ? magnitude : -magnitude) * quantization[component.QuantId][natural]);
                    }
                    WriteDctBlock(pixels, single ? x : x * component.H + b % component.H,
                        single ? y : y * component.V + b / component.H);
                }
            }
        }
        reader.Finish();
        foreach (int ci in scan.ComponentIndices) state.DecodedComponents[ci] = true;
    }

    private static void DecodeArithmeticDc(ArithmeticReader reader, byte[] bins,
        ArithmeticConditioning conditioning, int table, ref int context, ref int previous) {
        int difference = 0, s = context;
        if (reader.Decode(ref bins[s]) != 0) {
            int sign = reader.Decode(ref bins[s + 1]);
            int magnitude = DecodeArithmeticMagnitude(reader, bins, s + 2 + sign, 20, 21);
            difference = sign == 0 ? magnitude : -magnitude;
            context = magnitude <= (1 << conditioning.Lower[table]) / 2 ? 0
                : (magnitude > (1 << conditioning.Upper[table]) ? 12 : 4) + sign * 4;
        } else context = 0;
        // Prediction differences and accumulation use signed sixteen-bit wraparound.
        previous = unchecked((short)(previous + difference));
    }

    private static int DecodeArithmeticMagnitude(ArithmeticReader reader, byte[] bins, int first, int x1, int x2) {
        if (reader.Decode(ref bins[first]) == 0) return 1;
        int magnitude = 1, context = x1;
        if (reader.Decode(ref bins[x1]) != 0) {
            magnitude = 2; context = x2;
            while (reader.Decode(ref bins[context]) != 0) {
                magnitude <<= 1; context++;
                if (magnitude >= 32768) throw new FormatException("Arithmetic JPEG magnitude exceeds sixteen bits.");
            }
        }
        int value = magnitude;
        while ((magnitude >>= 1) != 0)
            if (reader.Decode(ref bins[context + 14]) != 0) value |= magnitude;
        return value + 1;
    }
}
