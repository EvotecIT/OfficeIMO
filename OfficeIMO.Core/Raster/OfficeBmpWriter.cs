using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Writes Windows V4 bitmap images with explicit straight-alpha bit masks.</summary>
internal static class OfficeBmpWriter {
    internal static void EncodeTo(OfficeRasterImage image, Stream output, OfficeRasterEncodingOptions options, CancellationToken token) {
        int size = OfficeRasterGuards.EnsureOutputBytes(122L + image.Width * (long)image.Height * 4, "Bitmap output exceeds the encoded-size limit.");
        var header = new byte[122]; header[0] = 66; header[1] = 77;
        Put(2, (uint)size); Put(10, 122); Put(14, 108); Put(18, (uint)image.Width); Put(22, (uint)image.Height);
        OfficeExifProfileCodec.Write(header, 26, 1, 2, true); OfficeExifProfileCodec.Write(header, 28, 32, 2, true);
        Put(30, 3); Put(34, (uint)(size - header.Length));
        if (options.WriteResolutionMetadata) { Put(38, checked((uint)Math.Round(options.DpiX / 0.0254D))); Put(42, checked((uint)Math.Round(options.DpiY / 0.0254D))); }
        Put(54, 0x00FF0000); Put(58, 0x0000FF00); Put(62, 0x000000FF); Put(66, 0xFF000000); Put(70, 0x73524742);
        output.Write(header, 0, header.Length); var row = new byte[image.Width * 4]; byte[] rgba = image.PixelBuffer;
        for (int y = image.Height - 1; y >= 0; y--) { token.ThrowIfCancellationRequested(); for (int x = 0; x < image.Width; x++) { int src = y * row.Length + x * 4; int dst = x * 4; row[dst] = rgba[src + 2]; row[dst + 1] = rgba[src + 1]; row[dst + 2] = rgba[src]; row[dst + 3] = rgba[src + 3]; } output.Write(row, 0, row.Length); }
        void Put(int at, uint value) => OfficeExifProfileCodec.Write(header, at, value, 4, true);
    }
}
