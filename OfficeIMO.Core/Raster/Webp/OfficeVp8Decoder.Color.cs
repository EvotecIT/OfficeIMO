// Adapted from CodeGlyphX, commit fc25e2fcf795d9c9a09b88708c47bdfeed5c446d.
// Copyright CodeGlyphX contributors. Apache-2.0; see THIRD-PARTY-NOTICES.md.
// OfficeIMO adaptation removes diagnostic scaffolding and adds bounded cancellation/resource handling.
namespace OfficeIMO.Drawing;

internal static partial class OfficeVp8Decoder {
    // Chroma samples are centered in each 2x2 luma cell. Interpolate at the
    // luma pixel centers, extending boundary samples rather than adding color.
    private static byte InterpolateDecodedChroma(byte[] plane, int stride, int width, int height, int x, int y) {
        int centerX = x >> 1;
        int centerY = y >> 1;
        int adjacentX = (x & 1) == 0 ? centerX - 1 : centerX + 1;
        int adjacentY = (y & 1) == 0 ? centerY - 1 : centerY + 1;
        int center = SampleDecodedChroma(plane, stride, width, height, centerX, centerY);
        int horizontal = SampleDecodedChroma(plane, stride, width, height, adjacentX, centerY);
        int vertical = SampleDecodedChroma(plane, stride, width, height, centerX, adjacentY);
        int diagonal = SampleDecodedChroma(plane, stride, width, height, adjacentX, adjacentY);
        return (byte)((9 * center + 3 * horizontal + 3 * vertical + diagonal + 8) >> 4);
    }
    private static byte SampleDecodedChroma(byte[] plane, int stride, int width, int height, int x, int y) {
        x = System.Math.Max(0, System.Math.Min(width - 1, x));
        y = System.Math.Max(0, System.Math.Min(height - 1, y));
        return plane[y * stride + x];
    }
}
