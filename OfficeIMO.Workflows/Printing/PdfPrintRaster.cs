using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>Pixel operations shared by preview and raster printer transports.</summary>
internal static class PdfPrintRaster {
    /// <summary>Converts RGBA pixels in place with the shared 77:150:29 luminance weights, one row at a time.</summary>
    internal static void ApplyGrayscale(OfficeRasterImage raster, CancellationToken token) =>
        ConvertRows(raster.PixelBuffer, raster.Width, raster.Height, grayscale: true, toBgra: false, token);

    /// <summary>Converts RGBA pixels in place to the BGRA order GDI expects, optionally applying grayscale in the same pass.</summary>
    internal static void ConvertRows(byte[] pixels, int width, int height, bool grayscale, bool toBgra, CancellationToken token) {
        int stride = checked(width * 4);
        if (pixels.Length < checked(stride * height)) throw new ArgumentException("The pixel buffer is smaller than its dimensions.", nameof(pixels));
        if (!grayscale && !toBgra) return;
        for (int y = 0; y < height; y++) {
            if ((y & 63) == 0) token.ThrowIfCancellationRequested();
            int end = (y + 1) * stride;
            for (int index = y * stride; index < end; index += 4) {
                byte r = pixels[index], b = pixels[index + 2];
                if (grayscale) {
                    byte gray = (byte)((r * 77 + pixels[index + 1] * 150 + b * 29 + 128) >> 8);
                    pixels[index] = gray; pixels[index + 1] = gray; pixels[index + 2] = gray;
                } else {
                    pixels[index] = b; pixels[index + 2] = r;
                }
            }
        }
    }

    /// <summary>Chooses a raster resolution no higher than the device or caller limit, reduced to stay within the pixel budget.</summary>
    internal static double ResolveDpi(double deviceDpi, double maximumDpi, double widthPoints, double heightPoints, long maximumPixels) {
        double dpi = Math.Min(deviceDpi > 0 ? deviceDpi : maximumDpi, maximumDpi);
        double squareInches = widthPoints / 72D * (heightPoints / 72D);
        double budgetDpi = Math.Sqrt(maximumPixels / squareInches);
        dpi = Math.Min(dpi, Math.Floor(budgetDpi));
        if (dpi < 36) throw new InvalidOperationException("The raster pixel budget is too small for this paper size.");
        return dpi;
    }
}
