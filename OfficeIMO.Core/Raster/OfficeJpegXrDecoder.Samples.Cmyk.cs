using System;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    private static byte[] FormatCmykRgba(FrameHeader frame, int[][] primary, FrameHeader? alphaFrame,
            int[][]? alpha, OfficeIccColorProfile? profile, CancellationToken cancellation) {
        if (profile == null || profile.ComponentCount != 4)
            throw new FormatException("JPEG-XR CMYK requires a four-component ICC profile.");
        int stride = frame.Width + frame.Left + frame.Right;
        int alphaStride = alphaFrame == null ? 0 : alphaFrame.Width + alphaFrame.Left + alphaFrame.Right;
        int maximum = frame.BitDepth == 2 ? 65535 : 255;
        var output = new byte[checked(frame.Width * frame.Height * 4)];
        var channels = new double[4];
        for (int y = 0; y < frame.Height; y++) {
            cancellation.ThrowIfCancellationRequested();
            for (int x = 0; x < frame.Width; x++) {
                if ((x & 4095) == 0) cancellation.ThrowIfCancellationRequested();
                int index = (y + frame.Top) * stride + x + frame.Left, target = (y * frame.Width + x) * 4;
                ReadCmykSamples(frame, primary, index, channels);
                if (!profile.TryConvert(channels, OfficeIccRenderingIntent.RelativeColorimetric, out var color))
                    throw new FormatException("JPEG-XR CMYK ICC conversion failed.");
                output[target] = color.R; output[target + 1] = color.G; output[target + 2] = color.B;
                int a = alpha == null ? maximum : ScaleSample(
                    alpha[0][(y + alphaFrame!.Top) * alphaStride + x + alphaFrame.Left],
                    frame.AlphaPlane ?? alphaFrame!.Primary, frame.BitDepth);
                output[target + 3] = ToByte(a, maximum);
            }
        }
        return output;
    }

    // T.832 Tables 186-188: transformed CMYK has half bias for C/M/Y and
    // negative half bias for K. CMYKDIRECT reorders YUVK and uses full bias.
    internal static void ReadCmykSamples(FrameHeader frame, int[][] samples, int index, double[] channels) {
        long cyan, magenta, yellow, black;
        if (frame.OutputColor == 5) {
            cyan = samples[1][index]; magenta = samples[2][index];
            yellow = samples[3][index]; black = samples[0][index];
        } else {
            black = samples[3][index] + ((long)samples[0][index] >> 1);
            magenta = black - samples[0][index] - ((long)samples[1][index] >> 1);
            cyan = samples[1][index] + magenta + ((long)samples[2][index] >> 1);
            yellow = cyan - samples[2][index];
            int shift = frame.BitDepth == 2 ? frame.Primary.ShiftBits : 0;
            long bias = shift >= 16 ? 0 : (frame.BitDepth == 2 ? 32768 : 128) >> shift;
            int scale = frame.Primary.Scaled ? 3 : 0;
            long colorAdjustment = ((bias >> 1) - bias) << scale;
            cyan += colorAdjustment; magenta += colorAdjustment; yellow += colorAdjustment;
            black -= (bias + (bias >> 1)) << scale;
        }
        double maximum = frame.BitDepth == 2 ? 65535D : 255D;
        channels[0] = ScaleSample(cyan, frame.Primary, frame.BitDepth) / maximum;
        channels[1] = ScaleSample(magenta, frame.Primary, frame.BitDepth) / maximum;
        channels[2] = ScaleSample(yellow, frame.Primary, frame.BitDepth) / maximum;
        channels[3] = ScaleSample(black, frame.Primary, frame.BitDepth) / maximum;
    }
}
