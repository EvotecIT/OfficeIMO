using System;

namespace OfficeIMO.Drawing;

/// <summary>Composes bounded Main-8 YUV420 or monochrome planes into owned straight-alpha RGBA.</summary>
/// <remarks>Converts channel values using CICP range/matrix metadata. ICC, transfer and gamut transforms are not applied.</remarks>
internal static class OfficeAvifColorConverter {
    internal static OfficeRasterImage Compose(OfficeAv1ReconstructedFrame frame, OfficeAvifColorDescription color,
        OfficeAv1ReconstructedFrame? alpha, OfficeRasterDecodeOptions options) {
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        if(frame.BitDepth!=8 || (alpha!=null && alpha.BitDepth!=8))
            throw new FormatException("AVIF high-bit-depth composition is not qualified.");
        if (!OfficeRasterGuards.TryEnsurePixelCount(frame.Width, frame.Height, options.MaximumDecodedPixels, out int pixels) ||
            (frame.PlaneCount != 1 && frame.PlaneCount != 3) ||
            (alpha != null && (alpha.Width != frame.Width || alpha.Height != frame.Height || alpha.PlaneCount != 1)) ||
            pixels > options.MaximumInspectionWorkPixels)
            throw new FormatException("Invalid or unbounded AVIF composition planes.");
        long storage = frame.StorageBytes + (alpha?.StorageBytes ?? 0) + (long)pixels * 4;
        if (options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes - storage)
            throw new FormatException("AVIF composition exceeds retained memory.");
        double kr, kb;
        switch (color.Matrix) {
            case 1: kr = .2126; kb = .0722; break;
            case 2: case 5: case 6: kr = .299; kb = .114; break;
            case 4: kr = .30; kb = .11; break;
            case 7: kr = .212; kb = .087; break;
            case 9: kr = .2627; kb = .0593; break;
            default: throw new NotSupportedException("Unsupported AVIF color matrix.");
        }
        double kg = 1 - kr - kb, yRange = color.FullRange ? 255 : 219, uvRange = color.FullRange ? 255 : 224;
        int biasY = color.FullRange ? 0 : 16;
        byte[] rgba = new byte[checked(pixels * 4)];
        for (int y = 0; y < frame.Height; y++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < frame.Width; x++) {
                double luma = (frame.Value(0, x, y) - biasY) / yRange, cb = 0, cr = 0;
                if (frame.PlaneCount == 3) {
                    cb = (Chroma(frame, 1, x, y) - 128) / uvRange;
                    cr = (Chroma(frame, 2, x, y) - 128) / uvRange;
                }
                int index = (y * frame.Width + x) * 4;
                rgba[index] = Channel(luma + 2 * (1 - kr) * cr);
                rgba[index + 1] = Channel(luma - 2 * (kr * (1 - kr) * cr + kb * (1 - kb) * cb) / kg);
                rgba[index + 2] = Channel(luma + 2 * (1 - kb) * cb);
                rgba[index + 3] = alpha == null ? (byte)255 : (byte)alpha.Value(0, x, y);
            }
        }
        options.CancellationToken.ThrowIfCancellationRequested();
        return OfficeRasterImage.FromOwnedRgba32(frame.Width, frame.Height, rgba);
    }

    private static double Chroma(OfficeAv1ReconstructedFrame frame, int plane, int x, int y) {
        int cx = x >> 1, cy = y >> 1, width = (frame.Width + 1) / 2, height = (frame.Height + 1) / 2;
        int ax = Math.Max(0, Math.Min(width - 1, cx + ((x & 1) == 0 ? -1 : 1)));
        int ay = Math.Max(0, Math.Min(height - 1, cy + ((y & 1) == 0 ? -1 : 1)));
        return (9 * frame.Value(plane, cx, cy) + 3 * frame.Value(plane, ax, cy) +
            3 * frame.Value(plane, cx, ay) + frame.Value(plane, ax, ay)) / 16.0;
    }

    private static byte Channel(double value) => (byte)Math.Max(0, Math.Min(255, Math.Floor(value * 255 + .5)));
}
