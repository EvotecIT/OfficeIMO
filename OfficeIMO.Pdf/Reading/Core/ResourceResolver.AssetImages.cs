using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    internal static int ReadImageInteger(PdfDictionary dictionary, string key, Dictionary<int, PdfIndirectObject>? objects) {
        if (!dictionary.Items.TryGetValue(key, out PdfObject? value)) return 0;
        PdfObject? resolved = value is PdfReference && objects is not null ? PdfObjectLookup.ResolveChain(objects, value) : value;
        return resolved is PdfNumber { Value: >= 1 and <= int.MaxValue } number ? (int)number.Value : 0;
    }

    internal static bool IsRgbaPng(byte[] png) => png.Length > 25 && png[24] == 8 && png[25] == 6;

    private static byte[] EnsureRgbaPng(byte[] png, int width, int height, int maximumBytes,
        CancellationToken cancellationToken) {
        if (IsRgbaPng(png)) return png;
        if (!PdfImageBufferLimits.TryGetScanlineBufferSize(width, height, 4, maximumBytes, out _, out int length) ||
            !OfficePngReader.TryDecode(png, cancellationToken, out OfficeRasterImage? raster) || raster is null ||
            raster.Width != width || raster.Height != height) return png;

        int stride = checked(width * 4);
        var scanlines = new byte[length];
        byte[] pixels = raster.PixelBuffer;
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(pixels, row * stride, scanlines, row * (stride + 1) + 1, stride);
        }
        return OfficePngWriter.EncodeScanlines(width, height, 8, 6, scanlines, OfficePngCompression.Fastest);
    }
}
