using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>Pixel operations shared by preview and raster printer transports.</summary>
internal static class PdfPrintRaster {
    /// <summary>Converts RGBA pixels in place with the shared 77:150:29 luminance weights, one row at a time.</summary>
    internal static void ApplyGrayscale(OfficeRasterImage raster, CancellationToken token) {
        byte[] pixels = raster.PixelBuffer;
        int stride = checked(raster.Width * 4);
        for (int y = 0; y < raster.Height; y++) {
            if ((y & 63) == 0) token.ThrowIfCancellationRequested();
            int end = (y + 1) * stride;
            for (int index = y * stride; index < end; index += 4) {
                byte gray = (byte)((pixels[index] * 77 + pixels[index + 1] * 150 + pixels[index + 2] * 29 + 128) >> 8);
                pixels[index] = gray; pixels[index + 1] = gray; pixels[index + 2] = gray;
            }
        }
    }

    /// <summary>
    /// Copies RGBA pixels into the rows a GDI device-independent bitmap expects: 24-bit BGR, or 8-bit gray for a gray sheet, each row
    /// padded to four bytes. Alpha is dropped, as a 32-bit GDI bitmap ignores it. Smaller rows make the spooler's copy of each sheet cheaper.
    /// </summary>
    internal static byte[] ToGdiRows(byte[] pixels, int width, int height, bool gray, CancellationToken token) {
        if (pixels.Length < checked(width * 4 * height)) throw new ArgumentException("The pixel buffer is smaller than its dimensions.", nameof(pixels));
        int stride = GdiStride(width, gray);
        var rows = new byte[checked(stride * height)];
        for (int y = 0; y < height; y++) {
            if ((y & 63) == 0) token.ThrowIfCancellationRequested();
            int from = y * width * 4, to = y * stride;
            for (int x = 0; x < width; x++, from += 4) {
                if (gray) {
                    rows[to++] = pixels[from];
                } else {
                    rows[to++] = pixels[from + 2]; rows[to++] = pixels[from + 1]; rows[to++] = pixels[from];
                }
            }
        }
        return rows;
    }

    /// <summary>Bytes per GDI bitmap row: 3 or 1 per pixel, padded to a four-byte boundary.</summary>
    internal static int GdiStride(int width, bool gray) => checked((width * (gray ? 1 : 3) + 3) & ~3);

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
