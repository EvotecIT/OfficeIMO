using System;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Drawing;

internal static partial class OfficeImageMetadataInspector {
    // Extraction is opt-in: ordinary metadata inspection need not inflate or retain profiles.
    // A null result with hasProfile=true means an unusable or over-budget profile.
    internal static byte[]? ReadIccProfile(byte[] data, OfficeImageFormat format,
        int maximumBytes, CancellationToken token, out bool hasProfile) {
        if (maximumBytes < 1 || maximumBytes > 4 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        token.ThrowIfCancellationRequested();
        hasProfile = false;
        if (data.Length > OfficeRasterGuards.MaximumEncodedBytes) return null;
        if (format == OfficeImageFormat.JpegXr) return ReadJpegXrIcc(data, maximumBytes, token, out hasProfile);
        if (format == OfficeImageFormat.Jpeg) {
            var snapshot = new OfficeImageMetadataSnapshot();
            InspectJpeg(data, snapshot, 0, token, maximumBytes);
            hasProfile = (snapshot.Kinds & OfficeImageMetadataKinds.Icc) != 0;
            return snapshot.Icc;
        }
        if (format == OfficeImageFormat.Png) {
            for (int offset = 8; offset <= data.Length - 12;) {
                token.ThrowIfCancellationRequested();
                int length = ReadBigEndian(data, offset);
                if (length < 0 || length > data.Length - 12 - offset) return null;
                if (ReadAscii(data, offset + 4, 4) == "iCCP") {
                    hasProfile = true;
                    int start = offset + 8, end = start + length;
                    int nameEnd = start;
                    while (nameEnd < end && nameEnd - start <= 79 && data[nameEnd] != 0) nameEnd++;
                    if (nameEnd == start || nameEnd - start > 79 || nameEnd > end - 3 || data[nameEnd + 1] != 0) return null;
                    int compressedLength = end - nameEnd - 2;
                    if (compressedLength > maximumBytes + 65536) return null;
                    var compressed = new byte[compressedLength];
                    Buffer.BlockCopy(data, nameEnd + 2, compressed, 0, compressedLength);
                    try { return OfficeZlibCodec.Decompress(compressed, maximumBytes, cancellationToken: token); }
                    catch (Exception ex) when (ex is FormatException || ex is InvalidDataException ||
                        ex is ArgumentException || ex is OfficeDecompressionSizeLimitException || ex is NotSupportedException) { return null; }
                }
                offset += length + 12;
            }
        } else if (format == OfficeImageFormat.Tiff && OfficeTiffCodec.IsTiff(data) && data.Length >= 8) {
            bool little = data[0] == (byte)'I';
            int ifd = ReadUInt32(data, 4, little);
            if (ifd < 8 || ifd > data.Length - 2) return null;
            int count = ReadUInt16(data, ifd, little);
            for (int index = 0; index < count; index++) {
                token.ThrowIfCancellationRequested();
                int entry = ifd + 2 + index * 12;
                if (entry > data.Length - 12) return null;
                if (ReadUInt16(data, entry, little) != 34675) continue;
                hasProfile = true;
                int length = ReadUInt32(data, entry + 4, little);
                int type = ReadUInt16(data, entry + 2, little);
                if ((type != 1 && type != 7) || length < 1 || length > maximumBytes) return null;
                int start = length <= 4 ? entry + 8 : ReadUInt32(data, entry + 8, little);
                if (start < 0 || start > data.Length - length) return null;
                return Slice(data, start, length, token);
            }
        }
        return null;
    }
}
