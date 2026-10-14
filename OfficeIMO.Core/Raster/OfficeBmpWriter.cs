using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Writes Windows V4 bitmap images with explicit straight-alpha bit masks.</summary>
internal static class OfficeBmpWriter {
    private const int HeaderLength = 122;
    private const int ConversionBufferLength = 16 * 1024;

    internal static int GetEncodedSize(int width, int height) {
        int pixels = OfficeRasterGuards.EnsureOutputPixels(width, height, "Bitmap dimensions exceed pixel limits.");
        return OfficeRasterGuards.EnsureOutputBytes(HeaderLength + pixels * 4L, "Bitmap output exceeds the encoded-size limit.");
    }

    internal static byte[] Encode(
        OfficeRasterImage image, OfficeRasterEncodingOptions options, OfficeImageExportEncodingBudget budget,
        CancellationToken token, long additionalRetainedManagedBytes) {
        token.ThrowIfCancellationRequested();
        int size = GetEncodedSize(image.Width, image.Height);
        EnsureRemainingBytes(budget, size);
        GetDensity(options, out uint dpiX, out uint dpiY);
        int bufferLength = Math.Min(checked(image.Width * 4), ConversionBufferLength);
        EnsureWorkingSet(image, additionalRetainedManagedBytes, size + 24L, bufferLength);
        // The returned array is the sole output backing store. A fixed stream never grows or copies it.
        var bytes = new byte[size];
        using var output = new MemoryStream(bytes, 0, bytes.Length, writable: true, publiclyVisible: true);
        using var guarded = new OfficeImageExportEncodingStream(output, budget, token);
        Write(image, guarded, size, bufferLength, dpiX, dpiY, token);
        token.ThrowIfCancellationRequested();
        return bytes;
    }

    internal static void EncodeTo(
        OfficeRasterImage image, Stream output, OfficeRasterEncodingOptions options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        OfficeRasterOutput.EnsureWritable(output);
        int size = GetEncodedSize(image.Width, image.Height);
        GetDensity(options, out uint dpiX, out uint dpiY);
        int bufferLength = Math.Min(checked(image.Width * 4), ConversionBufferLength);
        OfficeRasterOutput.EnsureImageWriteWorkingSet(image, output, size, bufferLength, HeaderLength,
            "Bitmap encoding exceeds the managed working-set limit.");
        Write(image, output, size, bufferLength, dpiX, dpiY, token);
    }

    private static void GetDensity(OfficeRasterEncodingOptions options, out uint x, out uint y) {
        x = y = 0U;
        if (!options.WriteResolutionMetadata) return;
        x = ToPixelsPerMeter(options.ResolvedDpiX);
        y = ToPixelsPerMeter(options.ResolvedDpiY);
    }

    private static uint ToPixelsPerMeter(double dpi) => checked((uint)Math.Round(
        OfficeRasterImageEncoder.NormalizeDpi(OfficeImageExportFormat.Bmp, dpi) / 0.0254D,
        MidpointRounding.AwayFromZero));

    private static void EnsureRemainingBytes(OfficeImageExportEncodingBudget budget, int size) {
        long remaining = budget.RemainingBytes;
        if (size <= remaining) return;
        long used = budget.MaximumBytes - remaining;
        long actual = used > long.MaxValue - size ? long.MaxValue : used + size;
        throw new OfficeImageExportBatchLimitException(nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes), actual, budget.MaximumBytes);
    }

    private static void EnsureWorkingSet(OfficeRasterImage image, long additionalRetainedBytes, long outputPeak, int bufferLength) {
        try {
            long peak = checked(image.PixelBuffer.LongLength + 24L + additionalRetainedBytes + outputPeak +
                HeaderLength + 24L + bufferLength + 24L);
            if (peak <= OfficeRasterGuards.MaximumDecodedBytes) return;
        } catch (OverflowException) { }
        throw new ArgumentException("Bitmap encoding exceeds the managed working-set limit.", nameof(image));
    }

    private static void Write(OfficeRasterImage image, Stream output, int size, int bufferLength, uint dpiX, uint dpiY, CancellationToken token) {
        var header = new byte[HeaderLength];
        header[0] = 66; header[1] = 77;
        Put(2, (uint)size); Put(10, HeaderLength); Put(14, 108); Put(18, (uint)image.Width); Put(22, (uint)image.Height);
        OfficeExifProfileCodec.Write(header, 26, 1, 2, true);
        OfficeExifProfileCodec.Write(header, 28, 32, 2, true);
        Put(30, 3); Put(34, (uint)(size - HeaderLength)); Put(38, dpiX); Put(42, dpiY);
        Put(54, 0x00FF0000); Put(58, 0x0000FF00); Put(62, 0x000000FF); Put(66, 0xFF000000); Put(70, 0x73524742);
        token.ThrowIfCancellationRequested();
        output.Write(header, 0, header.Length);
        var buffer = new byte[bufferLength];
        byte[] rgba = image.PixelBuffer;
        int rowLength = checked(image.Width * 4);
        for (int y = image.Height - 1; y >= 0; y--) {
            int rowStart = y * rowLength;
            for (int offset = 0; offset < rowLength; offset += bufferLength) {
                token.ThrowIfCancellationRequested();
                int count = Math.Min(bufferLength, rowLength - offset);
                for (int at = 0; at < count; at += 4) {
                    int source = rowStart + offset + at;
                    buffer[at] = rgba[source + 2]; buffer[at + 1] = rgba[source + 1];
                    buffer[at + 2] = rgba[source]; buffer[at + 3] = rgba[source + 3];
                }
                token.ThrowIfCancellationRequested();
                output.Write(buffer, 0, count);
            }
        }
        token.ThrowIfCancellationRequested();
        void Put(int at, uint value) => OfficeExifProfileCodec.Write(header, at, value, 4, true);
    }
}
