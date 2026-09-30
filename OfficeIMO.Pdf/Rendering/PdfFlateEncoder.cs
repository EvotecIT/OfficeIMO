using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Writes complete zlib streams, including the final DEFLATE block for empty input.</summary>
internal static class PdfFlateEncoder {
    internal static byte[] Compress(byte[] data, CancellationToken cancellationToken = default) {
        Guard.NotNull(data, nameof(data));
        cancellationToken.ThrowIfCancellationRequested();
        if (data.Length == 0) return new byte[] { 0x78, 0x9C, 0x03, 0x00, 0x00, 0x00, 0x00, 0x01 };
        const int chunkSize = 64 * 1024;
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            for (int offset = 0; offset < data.Length; offset += chunkSize) {
                cancellationToken.ThrowIfCancellationRequested();
                deflate.Write(data, offset, Math.Min(chunkSize, data.Length - offset));
            }
        }
        uint a = 1, b = 0;
        for (int index = 0; index < data.Length; index++) {
            if (index % chunkSize == 0) cancellationToken.ThrowIfCancellationRequested();
            a = (a + data[index]) % 65521;
            b = (b + a) % 65521;
        }
        uint adler = (b << 16) | a;
        output.WriteByte((byte)(adler >> 24));
        output.WriteByte((byte)(adler >> 16));
        output.WriteByte((byte)(adler >> 8));
        output.WriteByte((byte)adler);
        output.TryGetBuffer(out ArraySegment<byte> buffer);
        var result = new byte[checked((int)output.Length)];
        for (int offset = 0; offset < result.Length; offset += chunkSize) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(buffer.Array!, buffer.Offset + offset, result, offset, Math.Min(chunkSize, result.Length - offset));
        }
        return result;
    }
}
