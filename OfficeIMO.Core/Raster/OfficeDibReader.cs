using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Decodes Windows device-independent bitmap payloads that omit the BMP file header,
/// as commonly embedded in RTF <c>\dibitmap</c> pictures.
/// </summary>
public static class OfficeDibReader {
    private const int BitmapFileHeaderSize = 14;
    private const int BitmapInfoHeaderSize = 40;

    /// <summary>Attempts to decode an uncompressed 24-bit or 32-bit DIB into an RGBA raster.</summary>
    public static bool TryDecode(byte[]? dibBytes, out OfficeRasterImage? image) =>
        TryDecode(dibBytes, new OfficeRasterDecodeOptions(), out image);

    internal static bool TryDecode(
        byte[]? dibBytes,
        OfficeRasterDecodeOptions options,
        out OfficeRasterImage? image) {
        image = null;
        try {
            if (options == null) throw new ArgumentNullException(nameof(options));
            options.Validate();
            options.CancellationToken.ThrowIfCancellationRequested();
            if (dibBytes == null || dibBytes.Length < BitmapInfoHeaderSize ||
                dibBytes.Length > options.MaximumEncodedBytes) return false;
            int dibHeaderSize = ReadInt32LittleEndian(dibBytes, 0);
            if (dibHeaderSize < BitmapInfoHeaderSize || dibHeaderSize > dibBytes.Length) return false;

            int width = ReadInt32LittleEndian(dibBytes, 4);
            int signedHeight = ReadInt32LittleEndian(dibBytes, 8);
            if (width <= 0 || signedHeight == 0 ||
                !OfficeRasterImageDecoder.IsWithinPixelLimit(
                    width, Math.Abs(signedHeight), options.MaximumDecodedPixels)) return false;

            int bitsPerPixel = ReadUInt16LittleEndian(dibBytes, 14);
            int compression = ReadInt32LittleEndian(dibBytes, 16);
            if ((bitsPerPixel != 24 && bitsPerPixel != 32) || compression != 0) return false;

            int pixelOffset = checked(BitmapFileHeaderSize + dibHeaderSize);
            int bmpLength = checked(BitmapFileHeaderSize + dibBytes.Length);
            long retainedManagedBytes = checked(
                options.RetainedManagedBytes + dibBytes.LongLength + 24L);
            if (bmpLength > OfficeRasterGuards.MaximumEncodedBytes ||
                !OfficeBmpReader.IsDecodeWorkingSetWithinLimit(
                    bmpLength, width, Math.Abs(signedHeight), retainedManagedBytes)) return false;
            byte[] bmpBytes = new byte[bmpLength];
            bmpBytes[0] = (byte)'B';
            bmpBytes[1] = (byte)'M';
            WriteInt32LittleEndian(bmpBytes, 2, bmpBytes.Length);
            WriteInt32LittleEndian(bmpBytes, 10, pixelOffset);
            CopyWithCancellation(
                dibBytes, bmpBytes, BitmapFileHeaderSize, options.CancellationToken);
            return OfficeBmpReader.TryDecode(
                bmpBytes, options.CancellationToken, retainedManagedBytes, out image);
        } catch (OperationCanceledException) {
            image = null;
            throw;
        } catch {
            image = null;
            return false;
        }
    }

    private static int ReadInt32LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

    private static int ReadUInt16LittleEndian(byte[] bytes, int offset) =>
        bytes[offset] | (bytes[offset + 1] << 8);

    private static void WriteInt32LittleEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    private static void CopyWithCancellation(
        byte[] source,
        byte[] destination,
        int destinationOffset,
        CancellationToken cancellationToken) {
        const int chunkBytes = 64 * 1024;
        int sourceOffset = 0;
        while (sourceOffset < source.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = Math.Min(chunkBytes, source.Length - sourceOffset);
            Buffer.BlockCopy(source, sourceOffset, destination, destinationOffset + sourceOffset, count);
            sourceOffset += count;
        }
    }
}
