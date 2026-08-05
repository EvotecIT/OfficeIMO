using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Decodes Windows device-independent bitmap payloads that omit the BMP file header,
/// as commonly embedded in RTF <c>\dibitmap</c> pictures.
/// </summary>
public static class OfficeDibReader {
    private const int BitmapFileHeaderSize = 14;
    private const int BitmapInfoHeaderSize = 40;

    /// <summary>Attempts to decode an uncompressed 24-bit or 32-bit DIB into an RGBA raster.</summary>
    public static bool TryDecode(byte[]? dibBytes, out OfficeRasterImage? image) {
        image = null;
        try {
            if (dibBytes == null || dibBytes.Length < BitmapInfoHeaderSize) return false;
            int dibHeaderSize = ReadInt32LittleEndian(dibBytes, 0);
            if (dibHeaderSize < BitmapInfoHeaderSize || dibHeaderSize > dibBytes.Length) return false;

            int bitsPerPixel = ReadUInt16LittleEndian(dibBytes, 14);
            int compression = ReadInt32LittleEndian(dibBytes, 16);
            if ((bitsPerPixel != 24 && bitsPerPixel != 32) || compression != 0) return false;

            int pixelOffset = checked(BitmapFileHeaderSize + dibHeaderSize);
            byte[] bmpBytes = new byte[checked(BitmapFileHeaderSize + dibBytes.Length)];
            bmpBytes[0] = (byte)'B';
            bmpBytes[1] = (byte)'M';
            WriteInt32LittleEndian(bmpBytes, 2, bmpBytes.Length);
            WriteInt32LittleEndian(bmpBytes, 10, pixelOffset);
            Buffer.BlockCopy(dibBytes, 0, bmpBytes, BitmapFileHeaderSize, dibBytes.Length);
            return OfficeBmpReader.TryDecode(bmpBytes, out image);
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
}
