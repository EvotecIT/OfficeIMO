using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    // CIPA DC-007 MP Index offsets are relative to the TIFF endian field, except the first image (zero).
    private static bool HasCompleteMultiPictureJpeg(
        byte[] data, int tiffStart, int tiffLength, int primaryEnd, int fileEnd,
        CancellationToken cancellationToken, bool requireManagedFrame, bool validateMetadata,
        bool decodeImages, ref long remainingPixels) {
        bool littleEndian = data[tiffStart] == (byte)'I';
        uint Read32(int offset) => littleEndian
            ? ReadUInt32LittleEndian(data, offset) : ReadUInt32BigEndian(data, offset);
        int Read16(int offset) => littleEndian
            ? ReadUInt16LittleEndian(data, offset) : ReadUInt16BigEndian(data, offset);
        int directory = tiffStart + (int)Read32(tiffStart + 4);
        int count = Read16(directory);
        uint imageCount = 0;
        uint entryBytes = 0;
        uint entriesOffset = 0;
        bool hasVersion = false;
        for (int index = 0; index < count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int entry = directory + 2 + index * 12;
            int tag = Read16(entry);
            if (tag == 0xB000) {
                if (hasVersion || Read16(entry + 2) != 7 || Read32(entry + 4) != 4 ||
                    !HasJpegSegmentPrefix(data, entry + 8, 4, "0100")) return false;
                hasVersion = true;
            } else if (tag == 0xB001) {
                if (imageCount != 0 || Read16(entry + 2) != 4 || Read32(entry + 4) != 1) return false;
                imageCount = Read32(entry + 8);
            } else if (tag == 0xB002) {
                if (entryBytes != 0 || Read16(entry + 2) != 7) return false;
                entryBytes = Read32(entry + 4);
                entriesOffset = Read32(entry + 8);
            }
        }
        // APP2 bounds also bound the number of images; all arithmetic below is checked before casting.
        if (!hasVersion || imageCount == 0 || imageCount > (uint)tiffLength / 16 ||
            entryBytes != imageCount * 16 || entriesOffset < 8 || entriesOffset > tiffLength ||
            entryBytes > (uint)tiffLength - entriesOffset ||
            entriesOffset < directory - tiffStart + 2L + count * 12L + 4L &&
            entriesOffset + entryBytes > directory - tiffStart) return false;
        var ranges = new List<KeyValuePair<int, int>>((int)imageCount - 1);
        for (int index = 0; index < imageCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int entry = tiffStart + (int)entriesOffset + index * 16;
            uint attributes = Read32(entry);
            uint length = Read32(entry + 4);
            uint relativeOffset = Read32(entry + 8);
            if ((attributes & 0x07000000) != 0 || length < 12 ||
                Read16(entry + 12) > imageCount || Read16(entry + 14) > imageCount) return false;
            if (index == 0) {
                if (relativeOffset != 0 || length != primaryEnd) return false;
                continue;
            }
            long start = (long)tiffStart + relativeOffset;
            if (start < primaryEnd || start > fileEnd || length > fileEnd - start) return false;
            ranges.Add(new KeyValuePair<int, int>((int)start, (int)length));
        }
        ranges.Sort((left, right) => left.Key.CompareTo(right.Key));
        int next = primaryEnd;
        foreach (var range in ranges) {
            // Reject overlaps, undeclared payloads and truncation without copying each secondary image.
            if (range.Key != next || !HasCompleteJpegPayload(data, range.Key, range.Key + range.Value,
                cancellationToken, requireManagedFrame, validateMetadata, allowMultiPicture: false,
                decodeImages, ref remainingPixels)) return false;
            next += range.Value;
        }
        return next == fileEnd;
    }
}
