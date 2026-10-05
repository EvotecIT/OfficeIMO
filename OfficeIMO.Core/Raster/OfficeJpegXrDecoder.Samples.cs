using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal static int[][] ReconstructSamples(FrameHeader frame, PlaneHeader plane, DclpPlane dclp,
            int[]? highpass, CancellationToken cancellation) {
        int columns = (frame.Width + frame.Left + frame.Right) / 16, rows = (frame.Height + frame.Top + frame.Bottom) / 16;
        var output = new int[plane.Components][];
        var work = new long[16]; var block = new int[16];
        for (int c = 0; c < plane.Components; c++) {
            bool reduced = c != 0 && (plane.Color == 1 || plane.Color == 2);
            int blocksX = reduced ? 2 : 4, blocksY = reduced && plane.Color == 1 ? 2 : 4;
            int count = blocksX * blocksY, lowWidth = columns * blocksX, lowHeight = rows * blocksY;
            int width = lowWidth * 4, height = lowHeight * 4;
            output[c] = new int[checked(width * height)];
            for (int mb = 0; mb < columns * rows; mb++) {
                cancellation.ThrowIfCancellationRequested();
                int start = (mb * plane.Components + c) * 16;
                if (reduced) InverseReducedTransform(dclp.Coefficients, start, plane.Color, work);
                else InverseTransform(dclp.Coefficients, start, work);
                if (plane.Scaled && c != 0) for (int b = 0; b < count; b++)
                    dclp.Coefficients[start + b] = CheckedCoefficient((long)dclp.Coefficients[start + b] * 2);
            }
            if (frame.Overlap == 2) {
                var lowpass = new int[checked(lowWidth * lowHeight)];
                for (int y = 0; y < lowHeight; y++) {
                    cancellation.ThrowIfCancellationRequested();
                    for (int x = 0; x < lowWidth; x++)
                        lowpass[y * lowWidth + x] = dclp.Coefficients[((y / blocksY * columns + x / blocksX) * plane.Components + c) * 16 + y % blocksY * blocksX + x % blocksX];
                }
                if (reduced) ApplyReducedOverlap(frame, lowpass, lowWidth, lowHeight, blocksY, cancellation);
                else ApplyOverlap(frame, lowpass, lowWidth, lowHeight, 4, work, cancellation);
                for (int y = 0; y < lowHeight; y++) {
                    cancellation.ThrowIfCancellationRequested();
                    for (int x = 0; x < lowWidth; x++)
                        dclp.Coefficients[((y / blocksY * columns + x / blocksX) * plane.Components + c) * 16 + y % blocksY * blocksX + x % blocksX] = lowpass[y * lowWidth + x];
                }
            }
            for (int y = 0; y < rows; y++) for (int x = 0; x < columns; x++) {
                cancellation.ThrowIfCancellationRequested();
                int mb = y * columns + x, start = (mb * plane.Components + c) * 16;
                for (int b = 0; b < count; b++) {
                    if (highpass == null) Array.Clear(block, 0, block.Length);
                    else Array.Copy(highpass, (mb * plane.Components + c) * 256 + b * 16, block, 0, 16);
                    block[0] = dclp.Coefficients[start + b];
                    InverseTransform(block, 0, work);
                    int blockX = (x * blocksX + b % blocksX) * 4, blockY = (y * blocksY + b / blocksX) * 4;
                    for (int py = 0; py < 4; py++) for (int px = 0; px < 4; px++)
                        output[c][(blockY + py) * width + blockX + px] = block[py * 4 + px];
                }
            }
            if (frame.Overlap != 0)
                ApplyOverlap(frame, output[c], width, height, blocksX * 4, work, cancellation, blocksY * 4);
            if (reduced) output[c] = UpsampleChroma(output[c], width, height, plane, cancellation);
        }
        return output;
    }

    internal static byte[] FormatRgba(FrameHeader frame, int[][] primary, FrameHeader? alphaFrame, int[][]? alpha, bool premultiplied, OfficeIccColorProfile? colorProfile,
            CancellationToken cancellation) {
        int stride = frame.Width + frame.Left + frame.Right, alphaStride = alphaFrame == null ? 0
            : alphaFrame.Width + alphaFrame.Left + alphaFrame.Right;
        var output = new byte[checked(frame.Width * frame.Height * 4)];
        double[]? channels = colorProfile == null ? null : new double[colorProfile.ComponentCount];
        for (int y = 0; y < frame.Height; y++) {
            cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < frame.Width; x++) {
                if ((x & 4095) == 0) cancellation.ThrowIfCancellationRequested();
                int index = (y + frame.Top) * stride + x + frame.Left, target = (y * frame.Width + x) * 4;
                long red, green, blue;
                if (primary.Length == 1) red = green = blue = primary[0][index];
                else {
                    long chroma = -(long)primary[1][index];
                    green = primary[0][index] - (chroma >> 1);
                    red = chroma + green - (((long)primary[2][index] + 1) >> 1);
                    blue = primary[2][index] + red;
                }
                if (frame.BitDepth >= 3) {
                    double rExtended = ScaleExtendedSample(red, frame.Primary, frame.BitDepth);
                    double gExtended = ScaleExtendedSample(green, frame.Primary, frame.BitDepth);
                    double bExtended = ScaleExtendedSample(blue, frame.Primary, frame.BitDepth);
                    double aExtended = alpha == null ? 1D : ScaleExtendedSample(
                        alpha[0][(y + alphaFrame!.Top) * alphaStride + x + alphaFrame.Left],
                        frame.AlphaPlane ?? alphaFrame!.Primary, alphaFrame!.BitDepth);
                    WriteExtendedRgba(output, target, rExtended, gExtended, bExtended, aExtended, premultiplied, colorProfile, channels);
                    continue;
                }
                int maximum = frame.BitDepth == 2 ? 65535 : 255;
                int r = ScaleSample(red, frame.Primary, frame.BitDepth);
                int g = ScaleSample(green, frame.Primary, frame.BitDepth);
                int b = ScaleSample(blue, frame.Primary, frame.BitDepth);
                int a = maximum;
                if (alpha != null) {
                    int alphaIndex = (y + alphaFrame!.Top) * alphaStride + x + alphaFrame.Left;
                    a = ScaleSample(alpha[0][alphaIndex], frame.AlphaPlane ?? alphaFrame.Primary, alphaFrame.BitDepth);
                }
                // Unassociate and apply ICC at source precision before forming RGBA8.
                if (premultiplied) {
                    r = Unassociate(r, a, maximum); g = Unassociate(g, a, maximum); b = Unassociate(b, a, maximum);
                }
                if (colorProfile != null) {
                    channels![0] = r / (double)maximum;
                    if (channels.Length == 3) { channels[1] = g / (double)maximum; channels[2] = b / (double)maximum; }
                    if (!colorProfile.TryConvert(channels, OfficeIccRenderingIntent.RelativeColorimetric, out var color))
                        throw new FormatException("JPEG-XR ICC conversion failed.");
                    output[target] = color.R; output[target + 1] = color.G; output[target + 2] = color.B;
                } else {
                    output[target] = ToByte(r, maximum);
                    output[target + 1] = ToByte(g, maximum);
                    output[target + 2] = ToByte(b, maximum);
                }
                output[target + 3] = ToByte(a, maximum);
            }
        }
        return output;
    }

    private static int ScaleSample(long value, PlaneHeader plane, int bitDepth) {
        int maximum = bitDepth == 2 ? 65535 : 255;
        int shift = bitDepth == 2 ? plane.ShiftBits : 0;
        int bias = shift >= 16 ? 0 : ((maximum + 1) / 2) >> shift;
        value += (long)bias << (plane.Scaled ? 3 : 0);
        if (plane.Scaled) value = (value + (bitDepth == 2 ? 4 : 3)) >> 3;
        if (value <= 0) return 0;
        // Saturate before shifting, including declared shifts beyond the sample depth.
        if (shift >= 16 || value > (maximum >> shift)) return maximum;
        return (int)(value << shift);
    }

    private static int Unassociate(int value, int alpha, int maximum) =>
        alpha == 0 ? 0 : (int)Math.Min(maximum, ((long)value * maximum + alpha / 2) / alpha);

    private static byte ToByte(int value, int maximum) =>
        (byte)(((long)value * 255 + maximum / 2) / maximum);
}
