using System;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Bounded indexed, truecolor, and grayscale TGA decoding, including RLE, and 32-bit TGA writing.</summary>
internal static class OfficeTgaCodec {
    internal static bool TryIdentify(byte[] bytes, out int width, out int height) {
        width = height = 0;
        if (bytes.Length < 18) return false; int type = bytes[2]; int kind = type >= 9 ? type - 8 : type;
        if (kind < 1 || kind > 3 || (bytes[17] & 0xC0) != 0) return false;
        int bits = bytes[16]; if (kind == 1 ? bytes[1] != 1 || bits != 8 && bits != 16 : bytes[1] != 0 || (kind == 3 ? bits != 8 && bits != 16 : bits != 15 && bits != 16 && bits != 24 && bits != 32)) return false;
        width = U16(bytes, 12); height = U16(bytes, 14); return OfficeRasterGuards.TryEnsurePixelCount(width, height, out _);
    }
    internal static bool TryDecode(byte[] bytes, OfficeRasterDecodeOptions options, out OfficeRasterImage? image) {
        image = null;
        try {
            if (bytes.Length < 18) return false;
            int type = bytes[2]; bool rle = type >= 9; int kind = rle ? type - 8 : type;
            if (kind < 1 || kind > 3 || (bytes[17] & 0xC0) != 0) return false;
            int width = U16(bytes, 12); int height = U16(bytes, 14); int bits = bytes[16]; int alphaBits = bytes[17] & 15;
            bool indexed = kind == 1; bool gray = kind == 3;
            if (indexed) { if (bytes[1] != 1 || bits != 8 && bits != 16) return false; }
            else { if (bytes[1] != 0 || (gray ? bits != 8 && bits != 16 : bits != 15 && bits != 16 && bits != 24 && bits != 32)) return false; }
            if (alphaBits != 0 && alphaBits != 1 && alphaBits != 8 || alphaBits == 8 && bits != 32 && !(gray && bits == 16)) return false;
            if (!OfficeRasterGuards.TryEnsurePixelCount(width, height, options.MaximumDecodedPixels, out int pixels) || bytes.LongLength + pixels * 4L + options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes) return false;
            int cursor = checked(18 + bytes[0]); if (cursor > bytes.Length) return false;
            int first = U16(bytes, 3); int paletteLength = U16(bytes, 5); int paletteBits = bytes[7]; OfficeColor[]? palette = null;
            if (indexed) {
                if (paletteLength == 0 || first + paletteLength > 65536 || paletteBits != 15 && paletteBits != 16 && paletteBits != 24 && paletteBits != 32) return false;
                palette = new OfficeColor[paletteLength]; for (int i = 0; i < palette.Length; i++) palette[i] = ReadColor(paletteBits, false, paletteBits == 32 ? 8 : paletteBits == 16 ? 1 : 0);
            } else if (first != 0 || paletteLength != 0 || paletteBits != 0) return false;
            var result = new OfficeRasterImage(width, height); int written = 0;
            while (written < pixels) {
                options.CancellationToken.ThrowIfCancellationRequested();
                int packet = rle ? ReadByte() : 0; int count = rle ? (packet & 127) + 1 : 1; if (count > pixels - written) return false;
                bool repeat = rle && (packet & 128) != 0; OfficeColor color = OfficeColor.Transparent;
                for (int index = 0; index < count; index++) {
                    if (!repeat || index == 0) {
                        if (indexed) { int entry = bits == 8 ? ReadByte() : ReadByte() | ReadByte() << 8; if (entry < first || entry >= first + paletteLength) return false; color = palette![entry - first]; }
                        else color = ReadColor(bits, gray, alphaBits);
                    }
                    int x = written % width; int y = written / width; if ((bytes[17] & 16) != 0) x = width - 1 - x; if ((bytes[17] & 32) == 0) y = height - 1 - y;
                    result.SetPixel(x, y, color); written++;
                }
            }
            image = result; return true;
            byte ReadByte() { if (cursor >= bytes.Length) throw new FormatException("Truncated TGA pixel data."); return bytes[cursor++]; }
            OfficeColor ReadColor(int depth, bool grayscale, int alpha) {
                if (grayscale) { byte value = ReadByte(); return OfficeColor.FromRgba(value, value, value, depth == 16 ? ReadByte() : (byte)255); }
                if (depth == 15 || depth == 16) { int packed = ReadByte() | ReadByte() << 8; return OfficeColor.FromRgba((byte)(((packed >> 10) & 31) * 255 / 31), (byte)(((packed >> 5) & 31) * 255 / 31), (byte)((packed & 31) * 255 / 31), alpha == 1 ? (packed & 32768) != 0 ? (byte)255 : (byte)0 : (byte)255); }
                byte blue = ReadByte(); byte green = ReadByte(); byte red = ReadByte(); byte opacity = depth == 32 ? ReadByte() : (byte)255; return OfficeColor.FromRgba(red, green, blue, alpha == 8 ? opacity : (byte)255);
            }
        } catch (OperationCanceledException) { throw; } catch (Exception ex) when (ex is FormatException || ex is ArgumentException || ex is OverflowException || ex is IndexOutOfRangeException) { return false; }
    }
    internal static void EncodeTo(OfficeRasterImage image, Stream output, CancellationToken token) {
        if (image.Width > ushort.MaxValue || image.Height > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(image), "TGA dimensions must not exceed 65535.");
        OfficeRasterGuards.EnsureOutputBytes(18L + image.Width * (long)image.Height * 4, "TGA output exceeds the encoded-size limit.");
        var header = new byte[18]; header[2] = 2; OfficeExifProfileCodec.Write(header, 12, (uint)image.Width, 2, true); OfficeExifProfileCodec.Write(header, 14, (uint)image.Height, 2, true); header[16] = 32; header[17] = 40; output.Write(header, 0, header.Length);
        var row = new byte[image.Width * 4]; byte[] rgba = image.PixelBuffer;
        for (int y = 0; y < image.Height; y++) { token.ThrowIfCancellationRequested(); for (int x = 0; x < image.Width; x++) { int at = y * row.Length + x * 4; int dst = x * 4; row[dst] = rgba[at + 2]; row[dst + 1] = rgba[at + 1]; row[dst + 2] = rgba[at]; row[dst + 3] = rgba[at + 3]; } output.Write(row, 0, row.Length); }
    }
    private static int U16(byte[] bytes, int at) => bytes[at] | bytes[at + 1] << 8;
}
