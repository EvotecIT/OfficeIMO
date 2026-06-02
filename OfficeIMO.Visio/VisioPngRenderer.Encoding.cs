using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static partial class VisioPngRenderer {

        private static byte[] EncodePngRgba(int width, int height, byte[] rgba) {
            byte[] scanlines = new byte[height * (1 + width * 4)];
            int source = 0;
            int target = 0;
            for (int y = 0; y < height; y++) {
                scanlines[target++] = 0;
                Buffer.BlockCopy(rgba, source, scanlines, target, width * 4);
                source += width * 4;
                target += width * 4;
            }

            using MemoryStream ms = new();
            ms.Write(PngSignature, 0, PngSignature.Length);
            byte[] ihdr = new byte[13];
            WriteBigEndianInt32(ihdr, 0, width);
            WriteBigEndianInt32(ihdr, 4, height);
            ihdr[8] = 8;
            ihdr[9] = 6;
            WriteChunk(ms, "IHDR", ihdr);
            WriteChunk(ms, "IDAT", DeflateZlib(scanlines));
            WriteChunk(ms, "IEND", Array.Empty<byte>());
            return ms.ToArray();
        }

        private static byte[] DeflateZlib(byte[] data) {
            using MemoryStream ms = new();
            ms.WriteByte(0x78);
            ms.WriteByte(0x9C);
            using (DeflateStream deflate = new(ms, CompressionLevel.Optimal, leaveOpen: true)) {
                deflate.Write(data, 0, data.Length);
            }

            uint adler = Adler32(data);
            ms.WriteByte((byte)((adler >> 24) & 0xFF));
            ms.WriteByte((byte)((adler >> 16) & 0xFF));
            ms.WriteByte((byte)((adler >> 8) & 0xFF));
            ms.WriteByte((byte)(adler & 0xFF));
            return ms.ToArray();
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

        private static void WriteChunk(Stream stream, string type, byte[] data) {
            byte[] typeBytes = Encoding.ASCII.GetBytes(type);
            byte[] length = new byte[4];
            WriteBigEndianInt32(length, 0, data.Length);
            stream.Write(length, 0, length.Length);
            stream.Write(typeBytes, 0, typeBytes.Length);
            stream.Write(data, 0, data.Length);

            uint crc = Crc32(typeBytes, data);
            byte[] crcBytes = new byte[4];
            WriteBigEndianInt32(crcBytes, 0, unchecked((int)crc));
            stream.Write(crcBytes, 0, crcBytes.Length);
        }

        private static uint Crc32(byte[] type, byte[] data) {
            uint crc = 0xFFFFFFFF;
            for (int i = 0; i < type.Length; i++) crc = UpdateCrc(crc, type[i]);
            for (int i = 0; i < data.Length; i++) crc = UpdateCrc(crc, data[i]);
            return crc ^ 0xFFFFFFFF;
        }

        private static uint UpdateCrc(uint crc, byte value) {
            crc ^= value;
            for (int i = 0; i < 8; i++) {
                crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
            }

            return crc;
        }

        private static void WriteBigEndianInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)((value >> 24) & 0xFF);
            bytes[offset + 1] = (byte)((value >> 16) & 0xFF);
            bytes[offset + 2] = (byte)((value >> 8) & 0xFF);
            bytes[offset + 3] = (byte)(value & 0xFF);
        }
    }
}
