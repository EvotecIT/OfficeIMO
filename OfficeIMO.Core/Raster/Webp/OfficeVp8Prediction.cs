// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// OfficeIMO's copy is licensed under the repository's MIT license by the original author.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
using System;

namespace OfficeIMO.Drawing;

/// <summary>VP8 intra prediction shared by encoding and decoding.</summary>
internal static class OfficeVp8Prediction {
    private const int ModeVPred = 1, ModeHPred = 2, ModeTMPred = 3;
    private const int BModeTMPred = 1, BModeVEPred = 2, BModeHEPred = 3, BModeLDPred = 4,
        BModeRDPred = 5, BModeVRPred = 6, BModeVLPred = 7, BModeHDPred = 8, BModeHUPred = 9;
    private static byte ClampToByte(int value) => (byte)Math.Max(0, Math.Min(255, value));
    private static byte GetPlaneSampleOrDefault(byte[] plane, int width, int height, int x, int y, byte fallback) =>
        (uint)x < (uint)width && (uint)y < (uint)height ? plane[y * width + x] : fallback;
    internal static void PredictBlock(byte[] plane, int width, int height, int x, int y, int size, int mode, byte[] predicted, OfficeVp8DecodeScratch scratch) {
        int[] top = scratch.BlockTop;
        int[] left = scratch.BlockLeft;
        int sum = 0, count = 0;
        for (int i = 0; i < size; i++) {
            top[i] = GetPlaneSampleOrDefault(plane, width, height, Math.Min(width - 1, x + i), y - 1, 127);
            left[i] = GetPlaneSampleOrDefault(plane, width, height, x - 1, Math.Min(height - 1, y + i), 129);
            if (y > 0) { sum += top[i]; count++; }
            if (x > 0) { sum += left[i]; count++; }
        }
        int dc = count == 0 ? 128 : (sum + count / 2) / count;
        int corner = GetPlaneSampleOrDefault(plane, width, height, x - 1, y - 1, y == 0 ? (byte)127 : (byte)129);
        for (int row = 0; row < size; row++) {
            for (int col = 0; col < size; col++) {
                int value = mode switch { ModeVPred => top[col], ModeHPred => left[row], ModeTMPred => left[row] + top[col] - corner, _ => dc };
                predicted[row * size + col] = ClampToByte(value);
            }
        }
    }

    internal static void PredictSubblock(byte[] plane, int width, int height, int x, int y, int mode, byte[] predicted, OfficeVp8DecodeScratch scratch) {
        int[] top = scratch.SubblockTop;
        int[] left = scratch.SubblockLeft;
        int corner = GetPlaneSampleOrDefault(plane, width, height, x - 1, y - 1, y == 0 ? (byte)127 : (byte)129);
        for (int i = 0; i < 8; i++) {
            int topRow = i >= 4 && x % 16 == 12 ? y / 16 * 16 - 1 : y - 1;
            top[i] = GetPlaneSampleOrDefault(plane, width, height, Math.Min(width - 1, x + i), topRow, 127);
            left[i] = GetPlaneSampleOrDefault(plane, width, height, x - 1, Math.Min(height - 1, y + Math.Min(3, i)), 129);
        }
        int dc = (top[0] + top[1] + top[2] + top[3] + left[0] + left[1] + left[2] + left[3] + 4) >> 3;
        for (int row = 0; row < 4; row++) {
            for (int col = 0; col < 4; col++) {
                int value = mode switch {
                    BModeTMPred => left[row] + top[col] - corner,
                    BModeVEPred => Average3(SampleEdge(top, col - 1, corner), top[col], top[col + 1]),
                    BModeHEPred => Average3(SampleEdge(left, row - 1, corner), left[row], left[Math.Min(3, row + 1)]),
                    BModeLDPred => Average3(top[col + row], top[col + row + 1], top[Math.Min(7, col + row + 2)]),
                    BModeRDPred => DownRight(top, left, corner, col - row),
                    BModeVRPred => VerticalRight(top, left, corner, col, row),
                    BModeVLPred => VerticalLeft(top, col, row),
                    BModeHDPred => VerticalRight(left, top, corner, row, col),
                    BModeHUPred => VerticalLeft(left, row, col),
                    _ => dc
                };
                predicted[row * 4 + col] = ClampToByte(value);
            }
        }
    }

    private static int Average2(int a, int b) => (a + b + 1) >> 1;
    private static int Average3(int a, int b, int c) => (a + 2 * b + c + 2) >> 2;
    private static int SampleEdge(int[] edge, int index, int corner) => index < 0 ? corner : edge[Math.Min(edge.Length - 1, index)];
    private static int DiagonalSample(int[] top, int[] left, int corner, int index) =>
        index == 0 ? corner : index > 0 ? top[Math.Min(7, index - 1)] : left[Math.Min(7, -index - 1)];
    private static int DownRight(int[] top, int[] left, int corner, int offset) =>
        Average3(DiagonalSample(top, left, corner, offset - 1), DiagonalSample(top, left, corner, offset), DiagonalSample(top, left, corner, offset + 1));
    private static int VerticalRight(int[] top, int[] left, int corner, int x, int y) {
        int z = 2 * x - y;
        if (z < 0) return DownRight(top, left, corner, z + 1);
        int index = z / 2;
        return (z & 1) == 0
            ? Average2(SampleEdge(top, index - 1, corner), top[index])
            : Average3(SampleEdge(top, index - 1, corner), top[index], top[index + 1]);
    }
    private static int VerticalLeft(int[] top, int x, int y) {
        if (x == 3 && y >= 2) return Average3(top[y + 2], top[y + 3], top[y + 4]);
        int z = 2 * x + y, index = z / 2;
        return (z & 1) == 0 ? Average2(top[index], top[index + 1]) : Average3(top[index], top[index + 1], top[Math.Min(7, index + 2)]);
    }
}
