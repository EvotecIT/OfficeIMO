using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegReader {
    private static void SampleYccToRgb(JpegFrame frame, BaselineComponentState[] states,
        int yIndex, int cbIndex, int crIndex, int x, int y, bool highQualityChroma,
        out byte r, out byte g, out byte b) {
        int midpoint = 1 << (frame.Precision - 1);
        int luma = SampleComponent(states, yIndex, x, y, frame.MaxH, frame.MaxV,
            midpoint, highQualityChroma, preserveRaw16: true);
        int cb = SampleComponent(states, cbIndex, x, y, frame.MaxH, frame.MaxV,
            midpoint, highQualityChroma, preserveRaw16: true);
        int cr = SampleComponent(states, crIndex, x, y, frame.MaxH, frame.MaxV,
            midpoint, highQualityChroma, preserveRaw16: true);
        if (frame.Precision == 8) {
            YccToRgb(luma, cb, cr, out r, out g, out b);
            return;
        }
        // Chroma is centered at 2^(precision-1), not at half the inclusive
        // sample maximum. Convert in native units before projection and clipping;
        // projecting a two-bit neutral chroma value first would turn 2 into 170.
        double scale = 255D / ((1 << frame.Precision) - 1);
        cb -= midpoint;
        cr -= midpoint;
        r = ProjectColorSample((luma + 1.402D * cr) * scale);
        g = ProjectColorSample((luma - .344136286D * cb - .714136286D * cr) * scale);
        b = ProjectColorSample((luma + 1.772D * cb) * scale);
    }

    private static byte ProjectColorSample(double value) =>
        (byte)Math.Max(0, Math.Min(255, Math.Floor(value + .5D)));
}
