using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

internal static class PdfPngTestImages {
    internal static byte[] CreateTwoFrameApng() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 6, 0, 0, 0
        });
        WritePngChunk(ms, "acTL", new byte[] {
            0, 0, 0, 2,
            0, 0, 0, 0
        });
        byte[] frameControl = {
            0, 0, 0, 0,
            0, 0, 0, 1,
            0, 0, 0, 1,
            0, 0, 0, 0,
            0, 0, 0, 0,
            0, 1,
            0, 100,
            0,
            0
        };
        WritePngChunk(ms, "fcTL", frameControl);
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] { 0, 255, 0, 0, 255 }));

        frameControl[3] = 1;
        WritePngChunk(ms, "fcTL", frameControl);
        byte[] secondFrame = BuildStoredZlib(new byte[] { 0, 0, 255, 0, 255 });
        byte[] frameData = new byte[secondFrame.Length + 4];
        frameData[3] = 2;
        Buffer.BlockCopy(secondFrame, 0, frameData, 4, secondFrame.Length);
        WritePngChunk(ms, "fdAT", frameData);
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] Create16BitRgbPng(bool includeTransparency = false) {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            16, 2, 0, 0, 0
        });
        if (includeTransparency) {
            WritePngChunk(ms, "tRNS", new byte[] {
                0x12, 0x34,
                0x80, 0x00,
                0xFF, 0xFF
            });
        }

        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] {
            0,
            0x12, 0x34,
            0x80, 0x00,
            0xFF, 0xFF
        }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] Create16BitRgbaPng() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            16, 6, 0, 0, 0
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] {
            0,
            0x12, 0x34,
            0x80, 0x00,
            0xFF, 0xFF,
            0x40, 0x00
        }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateRgbaPngWithInvalidTransparencyChunk() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 6, 0, 0, 0
        });
        WritePngChunk(ms, "tRNS", new byte[] { 0, 0, 0, 0, 0, 0 });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] {
            0,
            0xFF,
            0x00,
            0x00,
            0x80
        }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] Create16BitGrayscaleAlphaPng() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            16, 4, 0, 0, 0
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] {
            0,
            0x80, 0x00,
            0x40, 0x00
        }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateInterlacedRgbPng() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 8,
            0, 0, 0, 8,
            8, 2, 0, 0, 1
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(CreateAdam7Scanlines(3, WriteRgbPixel)));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateRgbPng(int width, int height) {
        if (width <= 0) {
            throw new ArgumentOutOfRangeException(nameof(width));
        }

        if (height <= 0) {
            throw new ArgumentOutOfRangeException(nameof(height));
        }

        using var ms = CreatePng();
        var header = new byte[] {
            0, 0, 0, 0,
            0, 0, 0, 0,
            8, 2, 0, 0, 0
        };
        WriteInt32BigEndian(header, 0, width);
        WriteInt32BigEndian(header, 4, height);
        WritePngChunk(ms, "IHDR", header);

        var scanlines = new byte[(1 + width * 3) * height];
        for (int y = 0; y < height; y++) {
            int rowStart = y * (1 + width * 3);
            scanlines[rowStart] = 0;
            for (int x = 0; x < width; x++) {
                WriteRgbPixel(scanlines, rowStart + 1 + x * 3, x, y);
            }
        }

        WritePngChunk(ms, "IDAT", BuildStoredZlib(scanlines));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateRgbPng(byte red, byte green, byte blue) {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] {
            0,
            red,
            green,
            blue
        }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateRgbaPng(byte red, byte green, byte blue, byte alpha) {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 6, 0, 0, 0
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] {
            0,
            red,
            green,
            blue,
            alpha
        }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateInterlacedRgbaPng() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 8,
            0, 0, 0, 8,
            8, 6, 0, 0, 1
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(CreateAdam7Scanlines(4, WriteRgbaPixel)));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateInterlacedIndexedPng() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 8,
            0, 0, 0, 8,
            8, 3, 0, 0, 1
        });
        WritePngChunk(ms, "PLTE", new byte[] {
            0x12, 0x34, 0x56,
            0x80, 0x40, 0x20,
            0xEE, 0xDD, 0xCC,
            0x05, 0x99, 0xF0
        });
        WritePngChunk(ms, "tRNS", new byte[] { 255, 128, 0, 255 });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(CreateAdam7Scanlines(1, WriteIndexedPixel)));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreateOversizedInterlacedGrayscalePng() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0x00, 0x01, 0x86, 0xA0,
            0x00, 0x00, 0x75, 0x30,
            8, 0, 0, 0, 1
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(Array.Empty<byte>()));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreatePngWithInvalidCrc() {
        byte[] png = Create16BitRgbPng();
        png[png.Length - 1] ^= 0xFF;
        return png;
    }

    internal static byte[] CreateRgbPngWithExtraDecodedScanlines() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0
        });
        WritePngChunk(ms, "IDAT", BuildStoredZlib(new byte[] {
            0,
            0xFF,
            0x00,
            0x00,
            0,
            0x00,
            0xFF,
            0x00
        }));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        return ms.ToArray();
    }

    internal static byte[] CreatePngWithOverflowingChunkLength() {
        using var ms = CreatePng();
        WritePngChunk(ms, "IHDR", new byte[] {
            0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x01,
            8, 2, 0, 0, 0
        });
        ms.WriteByte(0x7F);
        ms.WriteByte(0xFF);
        ms.WriteByte(0xFF);
        ms.WriteByte(0xFF);
        byte[] type = Encoding.ASCII.GetBytes("tEXt");
        ms.Write(type, 0, type.Length);
        ms.Write(new byte[4], 0, 4);
        return ms.ToArray();
    }

    internal static byte[] CreateInterlacedRgbExpectedScanlines() {
        return CreateExpectedScanlines(3, WriteRgbPixel);
    }

    internal static byte[] CreateInterlacedRgbaExpectedScanlines() {
        return CreateExpectedScanlines(4, WriteRgbaPixel);
    }

    internal static byte[] CreateInterlacedIndexedRgbaExpectedScanlines() {
        return CreateExpectedScanlines(4, WriteIndexedRgbaPixel);
    }

    private static byte[] CreateExpectedScanlines(int channels, Action<byte[], int, int, int> writePixel) {
        var scanlines = new byte[(1 + 8 * 3) * 8];
        if (channels != 3) {
            scanlines = new byte[(1 + 8 * channels) * 8];
        }

        for (int y = 0; y < 8; y++) {
            int rowStart = y * (1 + 8 * channels);
            scanlines[rowStart] = 0;
            for (int x = 0; x < 8; x++) {
                int pixel = rowStart + 1 + x * channels;
                writePixel(scanlines, pixel, x, y);
            }
        }

        return scanlines;
    }

    internal static int ReadPngColorType(byte[] bytes) {
        Assert.True(bytes.Length > 25);
        Assert.Equal((byte)'I', bytes[12]);
        Assert.Equal((byte)'H', bytes[13]);
        Assert.Equal((byte)'D', bytes[14]);
        Assert.Equal((byte)'R', bytes[15]);
        return bytes[25];
    }

    internal static int ReadPngInterlaceMethod(byte[] bytes) {
        Assert.True(bytes.Length > 28);
        Assert.Equal((byte)'I', bytes[12]);
        Assert.Equal((byte)'H', bytes[13]);
        Assert.Equal((byte)'D', bytes[14]);
        Assert.Equal((byte)'R', bytes[15]);
        return bytes[28];
    }

    internal static byte[] DecodeStoredPngIdat(byte[] bytes) {
        using var idat = new MemoryStream();
        int offset = 8;
        while (offset + 12 <= bytes.Length) {
            int length = ReadInt32BigEndian(bytes, offset);
            Assert.True(length >= 0);
            Assert.True(offset + 12 + length <= bytes.Length);
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            if (type == "IDAT") {
                idat.Write(bytes, offset + 8, length);
            }

            if (type == "IEND") {
                break;
            }

            offset += 12 + length;
        }

        byte[] compressed = idat.ToArray();
        Assert.True(compressed.Length >= 6);
        Assert.Equal(0x78, compressed[0]);
        using var decoded = new MemoryStream();
        int compressedOffset = 2;
        bool finalBlock;
        do {
            Assert.True(compressedOffset + 5 <= compressed.Length);
            byte header = compressed[compressedOffset++];
            finalBlock = (header & 1) != 0;
            Assert.Equal(0, (header >> 1) & 0x03);

            int length = compressed[compressedOffset] | (compressed[compressedOffset + 1] << 8);
            int nlen = compressed[compressedOffset + 2] | (compressed[compressedOffset + 3] << 8);
            compressedOffset += 4;
            Assert.Equal(0xFFFF, length ^ nlen);
            Assert.True(compressedOffset + length <= compressed.Length - 4);
            decoded.Write(compressed, compressedOffset, length);
            compressedOffset += length;
        } while (!finalBlock);

        return decoded.ToArray();
    }

    internal static byte[] DecodePngIdat(byte[] bytes) {
        using var idat = new MemoryStream();
        int offset = 8;
        while (offset + 12 <= bytes.Length) {
            int length = ReadInt32BigEndian(bytes, offset);
            Assert.True(length >= 0);
            Assert.True(offset + 12 + length <= bytes.Length);
            string type = Encoding.ASCII.GetString(bytes, offset + 4, 4);
            if (type == "IDAT") {
                idat.Write(bytes, offset + 8, length);
            }

            if (type == "IEND") {
                break;
            }

            offset += 12 + length;
        }

        byte[] compressed = idat.ToArray();
        Assert.True(compressed.Length >= 6);
        Assert.Equal(0x78, compressed[0]);
        using var source = new MemoryStream(compressed, 2, compressed.Length - 6);
        using var deflate = new DeflateStream(source, CompressionMode.Decompress);
        using var decoded = new MemoryStream();
        deflate.CopyTo(decoded);
        return decoded.ToArray();
    }

    private static byte[] CreateAdam7Scanlines(int channels, Action<byte[], int, int, int> writePixel) {
        var scanlines = new MemoryStream();
        int[] xStarts = { 0, 4, 0, 2, 0, 1, 0 };
        int[] yStarts = { 0, 0, 4, 0, 2, 0, 1 };
        int[] xSteps = { 8, 8, 4, 4, 2, 2, 1 };
        int[] ySteps = { 8, 8, 8, 4, 4, 2, 2 };
        for (int pass = 0; pass < xStarts.Length; pass++) {
            for (int y = yStarts[pass]; y < 8; y += ySteps[pass]) {
                scanlines.WriteByte(0);
                for (int x = xStarts[pass]; x < 8; x += xSteps[pass]) {
                    var pixel = new byte[channels];
                    writePixel(pixel, 0, x, y);
                    scanlines.Write(pixel, 0, pixel.Length);
                }
            }
        }

        return scanlines.ToArray();
    }

    private static void WriteRgbPixel(byte[] buffer, int offset, int x, int y) {
        buffer[offset] = (byte)(x * 31);
        buffer[offset + 1] = (byte)(y * 29);
        buffer[offset + 2] = (byte)((x + y) * 15);
    }

    private static void WriteRgbaPixel(byte[] buffer, int offset, int x, int y) {
        WriteRgbPixel(buffer, offset, x, y);
        buffer[offset + 3] = (byte)(255 - (x * 17 + y * 11));
    }

    private static void WriteIndexedPixel(byte[] buffer, int offset, int x, int y) {
        buffer[offset] = (byte)((x + y) % 4);
    }

    private static void WriteIndexedRgbaPixel(byte[] buffer, int offset, int x, int y) {
        switch ((x + y) % 4) {
            case 0:
                buffer[offset] = 0x12;
                buffer[offset + 1] = 0x34;
                buffer[offset + 2] = 0x56;
                buffer[offset + 3] = 255;
                break;
            case 1:
                buffer[offset] = 0x80;
                buffer[offset + 1] = 0x40;
                buffer[offset + 2] = 0x20;
                buffer[offset + 3] = 128;
                break;
            case 2:
                buffer[offset] = 0xEE;
                buffer[offset + 1] = 0xDD;
                buffer[offset + 2] = 0xCC;
                buffer[offset + 3] = 0;
                break;
            default:
                buffer[offset] = 0x05;
                buffer[offset + 1] = 0x99;
                buffer[offset + 2] = 0xF0;
                buffer[offset + 3] = 255;
                break;
        }
    }

    private static MemoryStream CreatePng() {
        var ms = new MemoryStream();
        byte[] signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        ms.Write(signature, 0, signature.Length);
        return ms;
    }

    private static void WritePngChunk(Stream stream, string type, byte[] data) {
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        var length = new byte[4];
        WriteInt32BigEndian(length, 0, data.Length);
        stream.Write(length, 0, length.Length);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        uint crc = ComputeCrc32(typeBytes, data);
        var crcBytes = new byte[4];
        WriteUInt32BigEndian(crcBytes, 0, crc);
        stream.Write(crcBytes, 0, crcBytes.Length);
    }

    private static byte[] BuildStoredZlib(byte[] scanline) {
        using var ms = new MemoryStream();
        ms.WriteByte(0x78);
        ms.WriteByte(0x01);
        ms.WriteByte(0x01);
        ms.WriteByte((byte)(scanline.Length & 0xFF));
        ms.WriteByte((byte)((scanline.Length >> 8) & 0xFF));
        int nlen = scanline.Length ^ 0xFFFF;
        ms.WriteByte((byte)(nlen & 0xFF));
        ms.WriteByte((byte)((nlen >> 8) & 0xFF));
        ms.Write(scanline, 0, scanline.Length);
        uint adler = Adler32(scanline);
        ms.WriteByte((byte)((adler >> 24) & 0xFF));
        ms.WriteByte((byte)((adler >> 16) & 0xFF));
        ms.WriteByte((byte)((adler >> 8) & 0xFF));
        ms.WriteByte((byte)(adler & 0xFF));
        return ms.ToArray();
    }

    private static int ReadInt32BigEndian(byte[] buffer, int offset) {
        return (buffer[offset] << 24) |
               (buffer[offset + 1] << 16) |
               (buffer[offset + 2] << 8) |
               buffer[offset + 3];
    }

    private static void WriteInt32BigEndian(byte[] buffer, int offset, int value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static void WriteUInt32BigEndian(byte[] buffer, int offset, uint value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private static uint ComputeCrc32(byte[] typeBytes, byte[] data) {
        uint crc = 0xFFFFFFFF;
        for (int i = 0; i < typeBytes.Length; i++) {
            crc = UpdateCrc32(crc, typeBytes[i]);
        }

        for (int i = 0; i < data.Length; i++) {
            crc = UpdateCrc32(crc, data[i]);
        }

        return crc ^ 0xFFFFFFFF;
    }

    private static uint UpdateCrc32(uint crc, byte value) {
        crc ^= value;
        for (int i = 0; i < 8; i++) {
            if ((crc & 1) != 0) {
                crc = (crc >> 1) ^ 0xEDB88320;
            } else {
                crc >>= 1;
            }
        }

        return crc;
    }
}
