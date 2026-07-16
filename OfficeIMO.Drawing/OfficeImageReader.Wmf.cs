using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private const int PlaceableWmfHeaderSizeBytes = 22;
    private const int StandardWmfHeaderSizeBytes = 18;
    private const int StandardWmfRecordHeaderSizeBytes = 6;

    private static bool TryReadWmf(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length >= PlaceableWmfHeaderSizeBytes &&
            ReadInt32LittleEndian(data, 0) == unchecked((int)0x9AC6CDD7)) {
            return TryReadPlaceableWmf(data, out info);
        }

        return TryReadStandardWmf(data, headerOffset: 0, out info);
    }

    private static bool TryReadPlaceableWmf(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (!HasValidPlaceableWmfChecksum(data)) {
            return false;
        }

        if (!TryReadStandardWmf(data, PlaceableWmfHeaderSizeBytes, out _)) {
            return false;
        }

        int left = ReadInt16LittleEndian(data, 6);
        int top = ReadInt16LittleEndian(data, 8);
        int right = ReadInt16LittleEndian(data, 10);
        int bottom = ReadInt16LittleEndian(data, 12);
        int unitsPerInch = ReadUInt16LittleEndian(data, 14);
        if (unitsPerInch <= 0) {
            return false;
        }

        int width = (int)Math.Round(Math.Abs(right - left) * 96.0 / unitsPerInch);
        int height = (int)Math.Round(Math.Abs(bottom - top) * 96.0 / unitsPerInch);
        info = new OfficeImageInfo(OfficeImageFormat.Wmf, width, height);
        return width > 0 && height > 0;
    }

    private static bool TryReadStandardWmf(byte[] data, int headerOffset, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (headerOffset < 0 ||
            data.Length - headerOffset < StandardWmfHeaderSizeBytes + StandardWmfRecordHeaderSizeBytes) {
            return false;
        }

        int type = ReadUInt16LittleEndian(data, headerOffset);
        int headerSizeWords = ReadUInt16LittleEndian(data, headerOffset + 2);
        int version = ReadUInt16LittleEndian(data, headerOffset + 4);
        uint declaredSizeWords = ReadUInt32LittleEndian(data, headerOffset + 6);
        uint maximumRecordSizeWords = ReadUInt32LittleEndian(data, headerOffset + 12);
        int parameterCount = ReadUInt16LittleEndian(data, headerOffset + 16);
        long declaredSizeBytes = declaredSizeWords * 2L;
        long declaredEnd = headerOffset + declaredSizeBytes;
        if ((type != 1 && type != 2) ||
            headerSizeWords != StandardWmfHeaderSizeBytes / 2 ||
            (version != 0x0100 && version != 0x0300) ||
            declaredEnd != data.LongLength ||
            maximumRecordSizeWords < 3U ||
            parameterCount != 0) {
            return false;
        }

        int offset = headerOffset + StandardWmfHeaderSizeBytes;
        uint largestRecordSizeWords = 0U;
        while (offset < data.Length) {
            if ((long)offset + StandardWmfRecordHeaderSizeBytes > declaredEnd) {
                return false;
            }

            uint recordSizeWords = ReadUInt32LittleEndian(data, offset);
            int function = ReadUInt16LittleEndian(data, offset + 4);
            long nextOffset = (long)offset + (recordSizeWords * 2L);
            if (recordSizeWords < 3U ||
                recordSizeWords > maximumRecordSizeWords ||
                nextOffset > declaredEnd) {
                return false;
            }

            largestRecordSizeWords = Math.Max(largestRecordSizeWords, recordSizeWords);

            if (function == 0) {
                if (recordSizeWords != 3U ||
                    nextOffset != declaredEnd ||
                    largestRecordSizeWords != maximumRecordSizeWords) {
                    return false;
                }

                info = new OfficeImageInfo(OfficeImageFormat.Wmf, 0, 0);
                return true;
            }

            offset = (int)nextOffset;
        }

        return false;
    }

    private static bool HasValidPlaceableWmfChecksum(byte[] data) {
        int checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= ReadUInt16LittleEndian(data, offset);
        }

        return checksum == ReadUInt16LittleEndian(data, 20);
    }
}
