using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Core.Internal.Cabinet;

/// <summary>Decodes one MSZIP CFDATA frame with the preceding folder history, as specified by MS-MCI and RFC 1951.</summary>
internal static class OfficeMsZipDecoder {
    internal const int WindowBytes = 32768;
    internal const int MaximumCompressedBytes = WindowBytes + 12;

    /// <summary>Writes exactly the declared frame size, retaining the existing output prefix as the MSZIP dictionary.</summary>
    internal static void Decompress(byte[] cabinet, int offset, int compressedLength,
        byte[] output, int outputOffset, int expandedLength, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (compressedLength < 2 || compressedLength > MaximumCompressedBytes ||
            expandedLength < 1 || expandedLength > WindowBytes ||
            cabinet[offset] != 0x43 || cabinet[offset + 1] != 0x4B) {
            throw new OfficeCabinetException("CAB_MSZIP_HEADER", "The CAB contains an invalid MSZIP frame header or size.", offset);
        }

        int dictionaryLength = Math.Min(outputOffset, WindowBytes);
        // A non-final RFC 1951 stored block establishes the preceding 32 KiB as
        // ordinary output history. The original raw MSZIP Deflate blocks follow
        // at a byte boundary and retain their own final-block marker.
        var composed = new byte[5 + dictionaryLength + compressedLength - 2];
        composed[0] = 0;
        composed[1] = (byte)dictionaryLength;
        composed[2] = (byte)(dictionaryLength >> 8);
        ushort complement = unchecked((ushort)~dictionaryLength);
        composed[3] = (byte)complement;
        composed[4] = (byte)(complement >> 8);
        Buffer.BlockCopy(output, outputOffset - dictionaryLength, composed, 5, dictionaryLength);
        Buffer.BlockCopy(cabinet, offset + 2, composed, 5 + dictionaryLength, compressedLength - 2);

        // DeflateStream alone accepts some truncated streams on supported runtimes.
        // The canonical validator requires BFINAL, exact framing and bounded output.
        if (!OfficeDeflateStreamValidator.TryValidateExact(composed, 0, composed.Length,
                dictionaryLength + expandedLength, cancellationToken)) {
            throw new OfficeCabinetException("CAB_MSZIP_CORRUPT", "The MSZIP frame has invalid or incomplete Deflate data.", offset);
        }

        try {
            using (var source = new MemoryStream(composed, writable: false))
            using (var inflater = new DeflateStream(source, CompressionMode.Decompress)) {
                var discard = new byte[Math.Min(dictionaryLength, 8192)];
                int remaining = dictionaryLength;
                while (remaining > 0) {
                    int count = Math.Min(remaining, discard.Length);
                    ReadExactly(inflater, discard, 0, count, cancellationToken);
                    remaining -= count;
                }
                ReadExactly(inflater, output, outputOffset, expandedLength, cancellationToken);
                cancellationToken.ThrowIfCancellationRequested();
                if (inflater.ReadByte() != -1) {
                    throw new OfficeCabinetException("CAB_MSZIP_SIZE", "The MSZIP frame exceeds its declared expanded size.", offset);
                }
            }
        } catch (InvalidDataException exception) {
            throw new OfficeCabinetException("CAB_MSZIP_CORRUPT", "The MSZIP frame could not be decoded.", offset, exception);
        }
    }

    private static void ReadExactly(Stream stream, byte[] output, int offset, int count, CancellationToken cancellationToken) {
        int end = offset + count;
        while (offset < end) {
            cancellationToken.ThrowIfCancellationRequested();
            int read = stream.Read(output, offset, Math.Min(end - offset, 8192));
            if (read == 0) throw new OfficeCabinetException("CAB_MSZIP_SIZE", "The MSZIP frame is shorter than its declared expanded size.");
            offset += read;
        }
    }
}
