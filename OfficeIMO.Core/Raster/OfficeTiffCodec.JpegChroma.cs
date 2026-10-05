using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // Separate JPEG components have reduced SOF dimensions and 1x1 JPEG sampling.
    // TIFF places chroma at group centers (1) or at the first luma sample (2).
    private static void CopyTiffJpegChroma(byte[] decoded, byte[] source, int plane, int width,
        int left, int top, int columns, int rows, int componentWidth, int componentHeight,
        int horizontal, int vertical, int positioning, OfficeRasterDecodeOptions options) {
        int visibleWidth = Math.Min(componentWidth, checked((int)(((long)columns + horizontal - 1) / horizontal)));
        int visibleHeight = Math.Min(componentHeight, checked((int)(((long)rows + vertical - 1) / vertical)));
        for (int y = 0; y < rows; y++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            double fy = Math.Max(0, Math.Min(visibleHeight - 1, positioning == 2 ? y / (double)vertical : (y + .5) / vertical - .5));
            int y0 = (int)fy, y1 = Math.Min(y0 + 1, visibleHeight - 1);
            double dy = fy - y0;
            for (int x = 0; x < columns; x++) {
                if ((x & 4095) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                double fx = Math.Max(0, Math.Min(visibleWidth - 1, positioning == 2 ? x / (double)horizontal : (x + .5) / horizontal - .5));
                int x0 = (int)fx, x1 = Math.Min(x0 + 1, visibleWidth - 1);
                double dx = fx - x0;
                double upper = decoded[y0 * componentWidth + x0] * (1 - dx) + decoded[y0 * componentWidth + x1] * dx;
                double lower = decoded[y1 * componentWidth + x0] * (1 - dx) + decoded[y1 * componentWidth + x1] * dx;
                source[((top + y) * width + left + x) * 3 + plane] = (byte)Math.Round(upper * (1 - dy) + lower * dy);
            }
        }
    }
    private static void ReconstructTiffJpegChroma(byte[] pixels, int width, int columns, int rows,
        int horizontal, int vertical, int positioning, OfficeRasterDecodeOptions options) {
        if (horizontal == 1 && vertical == 1) return;
        int cw = checked((int)(((long)columns + horizontal - 1) / horizontal));
        int ch = checked((int)(((long)rows + vertical - 1) / vertical));
        int length = OfficeRasterGuards.EnsureByteCount((long)cw * ch, "TIFF chroma plane exceeds the managed limit.");
        // JPEG supplied nearest-neighbor component samples. Retain their original
        // grids before writing interpolated bytes back into the interleaved buffer.
        var cb = new byte[length];
        var cr = new byte[length];
        for (int y = 0; y < ch; y++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < cw; x++) {
                if ((x & 4095) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                int source = ((y * vertical) * width + x * horizontal) * 3;
                cb[y * cw + x] = pixels[source + 1];
                cr[y * cw + x] = pixels[source + 2];
            }
        }
        CopyTiffJpegChroma(cb, pixels, 1, width, 0, 0, columns, rows, cw, ch, horizontal, vertical, positioning, options);
        CopyTiffJpegChroma(cr, pixels, 2, width, 0, 0, columns, rows, cw, ch, horizontal, vertical, positioning, options);
    }

}
