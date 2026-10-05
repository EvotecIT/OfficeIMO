using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeTiffCodec {
    // Separate JPEG components have reduced SOF dimensions and 1x1 JPEG sampling.
    // TIFF positions their chroma samples at the centers of the corresponding luma groups.
    private static void CopyTiffJpegChroma(byte[] decoded, byte[] source, int plane, int width,
        int left, int top, int columns, int rows, int componentWidth, int componentHeight,
        int horizontal, int vertical, OfficeRasterDecodeOptions options) {
        for (int y = 0; y < rows; y++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            double fy = Math.Max(0, Math.Min(componentHeight - 1, (y + .5) / vertical - .5));
            int y0 = (int)fy, y1 = Math.Min(y0 + 1, componentHeight - 1);
            double dy = fy - y0;
            for (int x = 0; x < columns; x++) {
                if ((x & 4095) == 0) options.CancellationToken.ThrowIfCancellationRequested();
                double fx = Math.Max(0, Math.Min(componentWidth - 1, (x + .5) / horizontal - .5));
                int x0 = (int)fx, x1 = Math.Min(x0 + 1, componentWidth - 1);
                double dx = fx - x0;
                double upper = decoded[y0 * componentWidth + x0] * (1 - dx) + decoded[y0 * componentWidth + x1] * dx;
                double lower = decoded[y1 * componentWidth + x0] * (1 - dx) + decoded[y1 * componentWidth + x1] * dx;
                source[((top + y) * width + left + x) * 3 + plane] = (byte)Math.Round(upper * (1 - dy) + lower * dy);
            }
        }
    }
}
