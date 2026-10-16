using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    private static byte[] EncodeRgbaMaterialized(
        int width, int height, byte[] rgba, OfficePngCompression compression,
        double? dpiX, double? dpiY, CancellationToken cancellationToken, bool requireRgbColorSpace = false, bool requireGrayColorSpace = false) {
        using var output = new MemoryStream();
        EncodeRgbaStreaming(width, height, rgba, output, compression, dpiX, dpiY,
            cancellationToken, requireRgbColorSpace: requireRgbColorSpace, requireGrayColorSpace: requireGrayColorSpace);
        return output.ToArray();
    }

}
