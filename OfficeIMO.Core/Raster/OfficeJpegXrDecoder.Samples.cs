using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal static int[][] ReconstructSamples(FrameHeader frame, PlaneHeader plane, DclpPlane dclp,
            int[]? highpass, CancellationToken cancellation) {
        int columns = (frame.Width + frame.Left + frame.Right) / 16, rows = (frame.Height + frame.Top + frame.Bottom) / 16;
        int width = checked(columns * 16), height = checked(rows * 16), pixels = checked(width * height);
        var output = new int[plane.Components][];
        for (int c = 0; c < output.Length; c++) output[c] = new int[pixels];
        var work = new long[16]; var block = new int[16];
        for (int mb = 0; mb < columns * rows; mb++) {
            cancellation.ThrowIfCancellationRequested();
            for (int c = 0; c < plane.Components; c++) {
                int start = (mb * plane.Components + c) * 16;
                InverseTransform(dclp.Coefficients, start, work);
                if (plane.Scaled && c != 0) for (int b = 0; b < 16; b++)
                    dclp.Coefficients[start + b] = CheckedCoefficient((long)dclp.Coefficients[start + b] * 2);
            }
        }
        if (frame.Overlap == 2) {
            int lowWidth = columns * 4, lowHeight = rows * 4;
            var lowpass = new int[checked(lowWidth * lowHeight)];
            for (int c = 0; c < plane.Components; c++) {
                for (int y = 0; y < lowHeight; y++) {
                    cancellation.ThrowIfCancellationRequested();
                    for (int x = 0; x < lowWidth; x++)
                        lowpass[y * lowWidth + x] = dclp.Coefficients[((y / 4 * columns + x / 4) * plane.Components + c) * 16 + y % 4 * 4 + x % 4];
                }
                ApplyOverlap(frame, lowpass, lowWidth, lowHeight, 4, work, cancellation);
                for (int y = 0; y < lowHeight; y++) {
                    cancellation.ThrowIfCancellationRequested();
                    for (int x = 0; x < lowWidth; x++)
                        dclp.Coefficients[((y / 4 * columns + x / 4) * plane.Components + c) * 16 + y % 4 * 4 + x % 4] = lowpass[y * lowWidth + x];
                }
            }
        }
        for (int y = 0; y < rows; y++) for (int x = 0; x < columns; x++) {
            cancellation.ThrowIfCancellationRequested();
            int mb = y * columns + x;
            for (int c = 0; c < plane.Components; c++) {
                int start = (mb * plane.Components + c) * 16;
                for (int b = 0; b < 16; b++) {
                    if (highpass == null) Array.Clear(block, 0, block.Length);
                    else Array.Copy(highpass, (mb * plane.Components + c) * 256 + b * 16, block, 0, 16);
                    block[0] = dclp.Coefficients[start + b];
                    InverseTransform(block, 0, work);
                    int blockX = x * 16 + b % 4 * 4, blockY = y * 16 + b / 4 * 4;
                    for (int py = 0; py < 4; py++) for (int px = 0; px < 4; px++)
                        output[c][(blockY + py) * width + blockX + px] = block[py * 4 + px];
                }
            }
        }
        if (frame.Overlap != 0) for (int c = 0; c < output.Length; c++)
            ApplyOverlap(frame, output[c], width, height, 16, work, cancellation);
        return output;
    }

    internal static byte[] FormatRgba(FrameHeader frame, int[][] primary, FrameHeader? alphaFrame, int[][]? alpha,
            CancellationToken cancellation) {
        int stride = frame.Width + frame.Left + frame.Right, alphaStride = alphaFrame == null ? 0
            : alphaFrame.Width + alphaFrame.Left + alphaFrame.Right;
        var output = new byte[checked(frame.Width * frame.Height * 4)];
        for (int y = 0; y < frame.Height; y++) {
            cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < frame.Width; x++) {
                int index = (y + frame.Top) * stride + x + frame.Left, target = (y * frame.Width + x) * 4;
                long red, green, blue;
                if (primary.Length == 1) red = green = blue = primary[0][index];
                else {
                    long chroma = -(long)primary[1][index];
                    green = primary[0][index] - (chroma >> 1);
                    red = chroma + green - (((long)primary[2][index] + 1) >> 1);
                    blue = primary[2][index] + red;
                }
                output[target] = ScaleSample(red, frame.Primary.Scaled);
                output[target + 1] = ScaleSample(green, frame.Primary.Scaled);
                output[target + 2] = ScaleSample(blue, frame.Primary.Scaled);
                if (alpha == null) output[target + 3] = 255;
                else {
                    int alphaIndex = (y + alphaFrame!.Top) * alphaStride + x + alphaFrame.Left;
                    output[target + 3] = ScaleSample(alpha[0][alphaIndex], frame.AlphaPlane?.Scaled ?? alphaFrame.Primary.Scaled);
                }
            }
        }
        return output;
    }

    private static byte ScaleSample(long value, bool scaled) {
        value += scaled ? 1024 : 128;
        if (scaled) value = (value + 3) >> 3;
        return (byte)Math.Max(0, Math.Min(255, value));
    }
}
