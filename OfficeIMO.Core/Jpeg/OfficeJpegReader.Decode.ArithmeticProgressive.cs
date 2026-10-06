using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    // T.81 Annex G: retain one approximation position per spectral coefficient.
    // Zero means unseen; other values store Al+1, independently for each component.
    private sealed class ArithmeticProgression {
        public readonly byte[][] Positions;
        public ArithmeticProgression(int components) {
            Positions = new byte[components][];
            for (int i = 0; i < components; i++) Positions[i] = new byte[64];
        }

        public void Validate(ScanHeader scan) {
            if (scan.Ss > scan.Se || scan.Se > 63 || scan.Ss == 0 && scan.Se != 0 ||
                scan.Ss != 0 && scan.ComponentIndices.Length != 1 || scan.Al > 13 || scan.Ah > 13 ||
                scan.Ah != 0 && scan.Ah != scan.Al + 1)
                throw new FormatException("Invalid progressive arithmetic JPEG scan.");
            foreach (int ci in scan.ComponentIndices) {
                if (scan.Ss != 0 && Positions[ci][0] == 0)
                    throw new FormatException("Progressive JPEG AC data precedes DC data.");
                for (int k = scan.Ss; k <= scan.Se; k++) {
                    int prior = Positions[ci][k];
                    if (scan.Ah == 0 ? prior != 0 : prior != scan.Ah + 1)
                        throw new FormatException("Invalid progressive JPEG approximation order.");
                }
            }
        }

        public void Complete(ScanHeader scan) {
            foreach (int ci in scan.ComponentIndices)
                for (int k = scan.Ss; k <= scan.Se; k++) Positions[ci][k] = (byte)(scan.Al + 1);
        }

        public void RequireDcComponents() {
            foreach (byte[] positions in Positions)
                if (positions[0] == 0) throw new FormatException("Progressive JPEG component has no DC scan.");
        }
    }

    private static void DecodeArithmeticProgressive(OfficeByteView data, ScanHeader scan, JpegFrame frame,
        ProgressiveState state, ArithmeticProgression progression, int[][] quantization,
        ArithmeticConditioning conditioning, int restartInterval, CancellationToken token) {
        byte[][] bins = new byte[16][];
        int[] context = new int[frame.ComponentCount];
        bool dc = scan.Ss == 0, first = scan.Ah == 0;
        foreach (int ci in scan.ComponentIndices) {
            Component component = frame.Components[ci];
            var pixels = state.Components[ci];
            if (dc && first) {
                if (quantization[component.QuantId] is null) throw new FormatException("Missing JPEG quantization table.");
                // A table may arrive between component scans. Latch it on the
                // component's first scan; subsequent redefinitions serve other components.
                pixels.Quantization = quantization[component.QuantId];
            }
            pixels.Component = component;
            pixels.PrevDc = 0;
            if (!dc || first) bins[dc ? component.DcTable : component.AcTable] ??= new byte[dc ? 49 : 245];
        }
        var reader = new ArithmeticReader(data, token);
        bool single = scan.ComponentIndices.Length == 1;
        Component leading = frame.Components[scan.ComponentIndices[0]];
        int columns = single ? GetNonInterleavedBlockCount(frame.Width, leading.H, frame.MaxH) : state.McuCols;
        int rows = single ? GetNonInterleavedBlockCount(frame.Height, leading.V, frame.MaxV) : state.McuRows;
        int mcu = 0;
        for (int y = 0; y < rows; y++) for (int x = 0; x < columns; x++, mcu++) {
            token.ThrowIfCancellationRequested();
            if (restartInterval > 0 && mcu > 0 && mcu % restartInterval == 0) {
                reader.Restart();
                foreach (byte[] area in bins) if (area != null) Array.Clear(area, 0, area.Length);
                Array.Clear(context, 0, context.Length);
                foreach (int ci in scan.ComponentIndices) state.Components[ci].PrevDc = 0;
            }
            foreach (int ci in scan.ComponentIndices) {
                var component = state.Components[ci];
                int blocks = single ? 1 : component.Component.H * component.Component.V;
                for (int b = 0; b < blocks; b++) {
                    int bx = single ? x : x * component.Component.H + b % component.Component.H;
                    int by = single ? y : y * component.Component.V + b / component.Component.H;
                    int at = (by * component.BlocksPerRow + bx) * 64;
                    if (dc) {
                        if (first) {
                            DecodeArithmeticDc(reader, bins[component.Component.DcTable], conditioning,
                                component.Component.DcTable, ref context[ci], ref component.PrevDc);
                            SetProgressiveCoefficient(component.Coeffs, at, component.PrevDc << scan.Al);
                        } else if (reader.DecodeFixed() != 0) {
                            // DC point transforms round toward negative infinity;
                            // refinement restores the bit in its signed representation.
                            component.Coeffs[at] = unchecked((short)((ushort)component.Coeffs[at] | (ushort)(1 << scan.Al)));
                        }
                    } else if (first) {
                        DecodeArithmeticAcFirst(reader, bins[component.Component.AcTable], component.Coeffs, at,
                            scan, conditioning.AcBoundary[component.Component.AcTable]);
                    } else DecodeArithmeticAcRefine(reader, bins[component.Component.AcTable], component.Coeffs, at, scan);
                }
            }
        }
        reader.Finish();
        progression.Complete(scan);
    }

    private static void DecodeArithmeticAcFirst(ArithmeticReader reader, byte[] bins, short[] coefficients,
        int offset, ScanHeader scan, int boundary) {
        for (int k = scan.Ss; k <= scan.Se; k++) {
            int at = 3 * (k - 1);
            if (reader.Decode(ref bins[at]) != 0) break;
            while (reader.Decode(ref bins[at + 1]) == 0) {
                if (++k > scan.Se) throw new FormatException("Arithmetic JPEG coefficient run exceeds its band.");
                at += 3;
            }
            int sign = reader.DecodeFixed();
            int magnitude = DecodeArithmeticMagnitude(reader, bins, at + 2, at + 2, k <= boundary ? 189 : 217);
            SetProgressiveCoefficient(coefficients, offset + ZigZag[k], (sign == 0 ? magnitude : -magnitude) << scan.Al);
        }
    }

    private static void DecodeArithmeticAcRefine(ArithmeticReader reader, byte[] bins, short[] coefficients,
        int offset, ScanHeader scan) {
        int last = scan.Se, delta = 1 << scan.Al;
        while (last >= scan.Ss && coefficients[offset + ZigZag[last]] == 0) last--;
        for (int k = scan.Ss; k <= scan.Se; k++) {
            int at = 3 * (k - 1);
            if (k > last && reader.Decode(ref bins[at]) != 0) break;
            while (true) {
                int index = offset + ZigZag[k], value = coefficients[index];
                if (value != 0) {
                    if (reader.Decode(ref bins[at + 2]) != 0)
                        SetProgressiveCoefficient(coefficients, index, value + (value > 0 ? delta : -delta));
                    break;
                }
                if (reader.Decode(ref bins[at + 1]) != 0) {
                    coefficients[index] = (short)(reader.DecodeFixed() == 0 ? delta : -delta);
                    break;
                }
                if (++k > scan.Se) throw new FormatException("Arithmetic JPEG refinement run exceeds its band.");
                at += 3;
            }
        }
    }
}
