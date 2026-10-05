using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal sealed class DclpPlane {
        internal int[] Coefficients = Array.Empty<int>();
        internal byte[] HighpassModes = Array.Empty<byte>();
    }

    // Both neighbour prediction and HP scan selection use quantized coefficients.
    // Dequantized coefficients feed the two inverse transform stages.
    internal static DclpPlane ReconstructDclp(FrameHeader frame, BandData dc, BandData? lp,
            bool alpha, CancellationToken cancellation) {
        PlaneHeader plane = alpha ? frame.AlphaPlane! : frame.Primary;
        int components = plane.Components, columns = (frame.Width + frame.Left + frame.Right) / 16;
        int rows = (frame.Height + frame.Top + frame.Bottom) / 16, macroblocks = checked(columns * rows);
        int[] rawDc = alpha ? dc.Alpha : dc.Primary;
        int[]? rawLp = lp == null ? null : alpha ? lp.Alpha : lp.Primary;
        int[]? lpIndices = lp == null ? null : alpha ? lp.AlphaQuantizerIndices : lp.QuantizerIndices;
        int[][][] dcQuant = alpha ? dc.AlphaQuantizers : dc.Quantizers;
        int[][][]? lpQuant = lp == null ? null : alpha ? lp.AlphaQuantizers : lp.Quantizers;
        int length = checked(macroblocks * components * 16);
        if ((long)length * 8 + rawDc.Length * 4L + (rawLp?.Length ?? 0) * 4L > 256 * 1024 * 1024)
            throw new FormatException("JPEG-XR prediction working set exceeds the managed limit.");
        int[] predicted = new int[length];
        if (rawLp != null) Array.Copy(rawLp, predicted, length);
        var result = new DclpPlane { Coefficients = new int[length], HighpassModes = new byte[macroblocks] };
        int tile = 0, top = 0;
        for (int tileY = 0; tileY < frame.TileHeights.Length; tileY++) {
            int left = 0, height = frame.TileHeights[tileY];
            for (int tileX = 0; tileX < frame.TileWidths.Length; tileX++, tile++) {
                int width = frame.TileWidths[tileX];
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                    cancellation.ThrowIfCancellationRequested();
                    int mb = (top + y) * columns + left + x;
                    ReconstructMacroblockDclp(plane, columns, mb, x == 0, y == 0, rawDc, predicted,
                        lpIndices, dcQuant[tile][0], lpIndices == null ? null : lpQuant![tile][lpIndices[mb]], result);
                }
                left += width;
            }
            top += height;
        }
        return result;
    }

    private static void ReconstructMacroblockDclp(PlaneHeader plane, int columns, int mb, bool leftEdge,
            bool topEdge, int[] rawDc, int[] predicted, int[]? lpIndices, int[] dcQuant, int[]? lpQuant, DclpPlane result) {
        int components = plane.Components, start = mb * components * 16;
        int leftStart = start - components * 16, topStart = start - columns * components * 16;
        int mode = leftEdge ? topEdge ? 3 : 1 : topEdge ? 0 : DcPredictionMode(predicted, plane, leftStart, topStart, topStart - components * 16);
        int lpMode = lpIndices == null ? 2 : mode == 0 && lpIndices[mb] == lpIndices[mb - 1] ? 0
            : mode == 1 && lpIndices[mb] == lpIndices[mb - columns] ? 1 : 2;
        for (int c = 0; c < components; c++) {
            int channel = start + c * 16, l = leftStart + c * 16, t = topStart + c * 16;
            long prediction = mode == 0 ? predicted[l] : mode == 1 ? predicted[t]
                : mode == 2 ? ((long)predicted[l] + predicted[t] + (c != 0 && (plane.Color == 1 || plane.Color == 2) ? 1 : 0)) >> 1 : 0;
            predicted[channel] = CheckedCoefficient(rawDc[mb * components + c] + prediction);
            if (c != 0 && plane.Color == 1) {
                if (lpMode == 0) AddPrediction(predicted, channel + 2, l + 2);
                else if (lpMode == 1) AddPrediction(predicted, channel + 1, t + 1);
            } else if (c != 0 && plane.Color == 2) {
                if (lpMode == 0) {
                    AddPrediction(predicted, channel + 4, l + 4);
                    AddPrediction(predicted, channel + 2, l + 2);
                    AddPrediction(predicted, channel + 6, l + 6);
                } else if (lpMode == 1) {
                    AddPrediction(predicted, channel + 4, t + 4);
                    AddPrediction(predicted, channel + 1, t + 5);
                    AddPrediction(predicted, channel + 5, channel + 1);
                } else if (mode == 1) AddPrediction(predicted, channel + 5, channel + 1);
            } else {
                if (lpMode == 0) for (int k = 4; k <= 12; k += 4) AddPrediction(predicted, channel + k, l + k);
                else if (lpMode == 1) for (int k = 1; k <= 3; k++) AddPrediction(predicted, channel + k, t + k);
            }
            int dcScale = QuantMap(dcQuant[c], c == 0 ? 1 : 0, plane.Scaled);
            int lpScale = lpQuant == null ? 1 : QuantMap(lpQuant[c], c == 0 ? 1 : 0, plane.Scaled);
            result.Coefficients[channel] = CheckedCoefficient((long)predicted[channel] * dcScale);
            for (int k = 1; k < 16; k++) result.Coefficients[channel + k] = CheckedCoefficient((long)predicted[channel + k] * lpScale);
        }
        result.HighpassModes[mb] = HighpassPredictionMode(predicted, start, plane);
    }

    private static void AddPrediction(int[] values, int target, int source) =>
        values[target] = CheckedCoefficient((long)values[target] + values[source]);

    private static int DcPredictionMode(int[] values, PlaneHeader plane, int left, int top, int diagonal) {
        int components = plane.Components;
        long horizontal = 0, vertical = 0;
        for (int c = 0; c < components; c++) {
            int offset = c * 16, weight = c == 0 && components == 3 ? (plane.Color == 1 ? 8 : plane.Color == 2 ? 4 : 2) : 1;
            horizontal += Math.Abs((long)values[diagonal + offset] - values[left + offset]) * weight;
            vertical += Math.Abs((long)values[diagonal + offset] - values[top + offset]) * weight;
        }
        return horizontal * 4 < vertical ? 1 : vertical * 4 < horizontal ? 0 : 2;
    }

    private static byte HighpassPredictionMode(int[] values, int start, PlaneHeader plane) {
        int components = plane.Components;
        long horizontal = Math.Abs((long)values[start + 1]) + Math.Abs((long)values[start + 2]) + Math.Abs((long)values[start + 3]);
        long vertical = Math.Abs((long)values[start + 4]) + Math.Abs((long)values[start + 8]) + Math.Abs((long)values[start + 12]);
        for (int c = 1; c < components; c++) {
            horizontal += Math.Abs((long)values[start + c * 16 + 1]);
            vertical += Math.Abs((long)values[start + c * 16 + (plane.Color == 1 || plane.Color == 2 ? 2 : 4)]);
            if (plane.Color == 2) {
                horizontal += Math.Abs((long)values[start + c * 16 + 5]);
                vertical += Math.Abs((long)values[start + c * 16 + 6]);
            }
        }
        return horizontal * 4 < vertical ? (byte)0 : vertical * 4 < horizontal ? (byte)1 : (byte)2;
    }

    private static int QuantMap(int value, int scaledShift, bool scaled) {
        if (value == 0) return 1;
        if (scaled) return value < 16 ? value << scaledShift : (16 + value % 16) << ((value >> 4) - 1 + scaledShift);
        return value < 32 ? (value + 3) >> 2 : value < 48 ? ((17 + value % 16) >> 1) << ((value >> 4) - 2)
            : (16 + value % 16) << ((value >> 4) - 3);
    }

    private static int CheckedCoefficient(long value) {
        if (value < int.MinValue || value > int.MaxValue) throw new FormatException("JPEG-XR coefficient exceeds the managed range.");
        return (int)value;
    }
}
