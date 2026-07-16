using System;
using System.IO;
using System.IO.Compression;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>Encodes and decodes RFC 1950 zlib streams with checksum validation.</summary>
    internal static class OfficeZlibCodec {
        internal static byte[] Compress(byte[] bytes) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            using var output = new MemoryStream();
            output.WriteByte(0x78);
            output.WriteByte(0x9C);
            using (var deflate = new DeflateStream(output,
                       CompressionLevel.Optimal, leaveOpen: true)) {
                deflate.Write(bytes, 0, bytes.Length);
            }
            uint checksum = Adler32(bytes);
            output.WriteByte(unchecked((byte)(checksum >> 24)));
            output.WriteByte(unchecked((byte)(checksum >> 16)));
            output.WriteByte(unchecked((byte)(checksum >> 8)));
            output.WriteByte(unchecked((byte)checksum));
            return output.ToArray();
        }

        internal static byte[] Decompress(byte[] bytes, int maximumOutputBytes,
            int? expectedOutputBytes = null) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (maximumOutputBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));
            if (bytes.Length < 6) throw new InvalidDataException("The zlib stream is truncated.");

            int compressionMethodAndInfo = bytes[0];
            int flags = bytes[1];
            if ((compressionMethodAndInfo & 0x0F) != 8
                || (compressionMethodAndInfo >> 4) > 7
                || ((compressionMethodAndInfo << 8) + flags) % 31 != 0) {
                throw new InvalidDataException("The zlib stream header is invalid.");
            }
            if ((flags & 0x20) != 0) {
                throw new NotSupportedException("Preset-dictionary zlib streams are not supported.");
            }

            using var source = new MemoryStream(bytes, 2, bytes.Length - 6, writable: false);
            using var deflate = new DeflateStream(source, CompressionMode.Decompress);
            using var output = new MemoryStream(expectedOutputBytes.GetValueOrDefault());
            var buffer = new byte[8192];
            while (true) {
                int read = deflate.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (output.Length > maximumOutputBytes - read) {
                    throw new InvalidDataException(
                        $"The decompressed zlib stream exceeds {maximumOutputBytes} bytes.");
                }
                output.Write(buffer, 0, read);
            }

            byte[] result = output.ToArray();
            if (expectedOutputBytes.HasValue && result.Length != expectedOutputBytes.Value) {
                throw new InvalidDataException(
                    $"The zlib stream expanded to {result.Length} bytes instead of {expectedOutputBytes.Value} bytes.");
            }
            uint expectedChecksum = ReadBigEndianUInt32(bytes, bytes.Length - 4);
            if (Adler32(result) != expectedChecksum) {
                throw new InvalidDataException("The zlib stream Adler-32 checksum is invalid.");
            }
            return result;
        }

        private static uint Adler32(byte[] data) {
            const uint Modulus = 65521;
            uint a = 1;
            uint b = 0;
            for (int index = 0; index < data.Length; index++) {
                a = (a + data[index]) % Modulus;
                b = (b + a) % Modulus;
            }
            return (b << 16) | a;
        }

        private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
            unchecked(((uint)bytes[offset] << 24)
                | ((uint)bytes[offset + 1] << 16)
                | ((uint)bytes[offset + 2] << 8)
                | bytes[offset + 3]);
    }
}
