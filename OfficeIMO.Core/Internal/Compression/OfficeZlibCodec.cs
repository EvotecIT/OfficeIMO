using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Core.Internal {
    /// <summary>Encodes and decodes RFC 1950 zlib streams with checksum validation.</summary>
    internal static class OfficeZlibCodec {
        internal static byte[] Compress(byte[] bytes) {
            return Compress(bytes, CancellationToken.None);
        }

        internal static byte[] Compress(byte[] bytes, CancellationToken cancellationToken) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            cancellationToken.ThrowIfCancellationRequested();
            using var output = new MemoryStream();
            output.WriteByte(0x78);
            output.WriteByte(0x9C);
            using (var deflate = new DeflateStream(output,
                       CompressionLevel.Optimal, leaveOpen: true)) {
                const int chunkSize = 64 * 1024;
                for (int offset = 0; offset < bytes.Length; offset += chunkSize) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int count = Math.Min(chunkSize, bytes.Length - offset);
                    deflate.Write(bytes, offset, count);
                }
            }
            uint checksum = Adler32(bytes, cancellationToken);
            output.WriteByte(unchecked((byte)(checksum >> 24)));
            output.WriteByte(unchecked((byte)(checksum >> 16)));
            output.WriteByte(unchecked((byte)(checksum >> 8)));
            output.WriteByte(unchecked((byte)checksum));
            return output.ToArray();
        }

        internal static byte[] Decompress(byte[] bytes, int maximumOutputBytes,
            int? expectedOutputBytes = null,
            CancellationToken cancellationToken = default) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (maximumOutputBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));
            if (expectedOutputBytes.HasValue && expectedOutputBytes.Value < 0) {
                throw new ArgumentOutOfRangeException(nameof(expectedOutputBytes));
            }
            if (expectedOutputBytes.HasValue && expectedOutputBytes.Value > maximumOutputBytes) {
                throw new OfficeDecompressionSizeLimitException(
                    $"The decompressed zlib stream exceeds {maximumOutputBytes} bytes.");
            }
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
            cancellationToken.ThrowIfCancellationRequested();
            int validationOutputLimit = expectedOutputBytes ?? maximumOutputBytes;
            if (!OfficeDeflateStreamValidator.TryValidateExact(
                    bytes, 2, bytes.Length - 6, validationOutputLimit,
                    out bool outputLimitExceeded, cancellationToken)) {
                if (outputLimitExceeded) {
                    if (expectedOutputBytes.HasValue && expectedOutputBytes.Value < maximumOutputBytes) {
                        throw new InvalidDataException(
                            $"The zlib stream expanded beyond the expected {expectedOutputBytes} bytes.");
                    }
                    throw new OfficeDecompressionSizeLimitException(
                        $"The decompressed zlib stream exceeds {maximumOutputBytes} bytes.");
                }
                throw new InvalidDataException("The zlib stream contains an invalid or trailing Deflate payload.");
            }

            using var source = new MemoryStream(bytes, 2, bytes.Length - 6, writable: false);
            using var deflate = new DeflateStream(source, CompressionMode.Decompress);
            byte[] result = expectedOutputBytes.HasValue
                ? DecompressExact(deflate, expectedOutputBytes.Value, maximumOutputBytes, cancellationToken)
                : DecompressBounded(deflate, maximumOutputBytes, cancellationToken);
            uint expectedChecksum = ReadBigEndianUInt32(bytes, bytes.Length - 4);
            if (Adler32(result, cancellationToken) != expectedChecksum) {
                throw new InvalidDataException("The zlib stream Adler-32 checksum is invalid.");
            }
            return result;
        }

        private static uint Adler32(byte[] data, CancellationToken cancellationToken = default) {
            const uint Modulus = 65521;
            const int MaximumChunk = 5552;
            uint a = 1;
            uint b = 0;
            int offset = 0;
            while (offset < data.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int end = Math.Min(offset + MaximumChunk, data.Length);
                while (offset < end) {
                    a += data[offset++];
                    b += a;
                }
                a %= Modulus;
                b %= Modulus;
            }
            return (b << 16) | a;
        }

        private static byte[] DecompressExact(
            DeflateStream deflate,
            int expectedOutputBytes,
            int maximumOutputBytes,
            CancellationToken cancellationToken) {
            if (expectedOutputBytes < 0) throw new ArgumentOutOfRangeException(nameof(expectedOutputBytes));
            if (expectedOutputBytes > maximumOutputBytes) {
                throw new OfficeDecompressionSizeLimitException(
                    $"The decompressed zlib stream exceeds {maximumOutputBytes} bytes.");
            }

            var result = new byte[expectedOutputBytes];
            int offset = 0;
            while (offset < result.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = deflate.Read(result, offset, result.Length - offset);
                if (read == 0) {
                    throw new InvalidDataException(
                        $"The zlib stream expanded to {offset} bytes instead of {expectedOutputBytes} bytes.");
                }
                offset += read;
            }
            if (deflate.ReadByte() != -1) {
                throw new InvalidDataException(
                    $"The zlib stream expanded beyond the expected {expectedOutputBytes} bytes.");
            }
            return result;
        }

        private static byte[] DecompressBounded(
            DeflateStream deflate,
            int maximumOutputBytes,
            CancellationToken cancellationToken) {
            using var output = new MemoryStream();
            var buffer = new byte[8192];
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = deflate.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                if (output.Length > maximumOutputBytes - read) {
                    throw new OfficeDecompressionSizeLimitException(
                        $"The decompressed zlib stream exceeds {maximumOutputBytes} bytes.");
                }
                output.Write(buffer, 0, read);
            }
            return output.ToArray();
        }

        private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
            unchecked(((uint)bytes[offset] << 24)
                | ((uint)bytes[offset + 1] << 16)
                | ((uint)bytes[offset + 2] << 8)
                | bytes[offset + 3]);
    }
}
