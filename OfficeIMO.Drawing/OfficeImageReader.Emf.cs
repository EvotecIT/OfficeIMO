using System;

namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private const int EmfHeaderMinimumSizeBytes = 88;
    private const int EmfRecordHeaderSizeBytes = 8;
    private const int EmfEofMinimumSizeBytes = 20;

    private static bool TryReadEmf(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < EmfHeaderMinimumSizeBytes) {
            return false;
        }

        int recordType = ReadInt32LittleEndian(data, 0);
        int headerSize = ReadInt32LittleEndian(data, 4);
        int signature = ReadInt32LittleEndian(data, 40);
        uint declaredBytes = ReadUInt32LittleEndian(data, 48);
        uint declaredRecords = ReadUInt32LittleEndian(data, 52);
        if (recordType != 1 ||
            headerSize < EmfHeaderMinimumSizeBytes ||
            (headerSize & 3) != 0 ||
            headerSize > data.Length ||
            signature != 0x464D4520 ||
            declaredBytes != data.LongLength ||
            declaredRecords < 2 ||
            !HasValidEmfRecordStream(data, headerSize, declaredRecords)) {
            return false;
        }

        int frameLeft = ReadInt32LittleEndian(data, 24);
        int frameTop = ReadInt32LittleEndian(data, 28);
        int frameRight = ReadInt32LittleEndian(data, 32);
        int frameBottom = ReadInt32LittleEndian(data, 36);
        int deviceWidth = ReadInt32LittleEndian(data, 72);
        int deviceHeight = ReadInt32LittleEndian(data, 76);
        int millimetersWidth = ReadInt32LittleEndian(data, 80);
        int millimetersHeight = ReadInt32LittleEndian(data, 84);

        double dpiX = millimetersWidth > 0 && deviceWidth > 0 ? deviceWidth * 25.4 / millimetersWidth : 96.0;
        double dpiY = millimetersHeight > 0 && deviceHeight > 0 ? deviceHeight * 25.4 / millimetersHeight : 96.0;
        bool hasFrameWidth = TryConvertPixelDimension(
            Math.Abs((long)frameRight - frameLeft) / 2540.0 * dpiX,
            out int width);
        bool hasFrameHeight = TryConvertPixelDimension(
            Math.Abs((long)frameBottom - frameTop) / 2540.0 * dpiY,
            out int height);

        if (!hasFrameWidth || !hasFrameHeight) {
            int boundsLeft = ReadInt32LittleEndian(data, 8);
            int boundsTop = ReadInt32LittleEndian(data, 12);
            int boundsRight = ReadInt32LittleEndian(data, 16);
            int boundsBottom = ReadInt32LittleEndian(data, 20);
            if (!TryConvertPixelDimension(Math.Abs((long)boundsRight - boundsLeft), out width) ||
                !TryConvertPixelDimension(Math.Abs((long)boundsBottom - boundsTop), out height)) {
                return false;
            }
        }

        info = new OfficeImageInfo(OfficeImageFormat.Emf, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static bool HasValidEmfRecordStream(byte[] data, int headerSize, uint declaredRecords) {
        int recordOffset = headerSize;
        uint recordCount = 1;
        while (recordOffset < data.Length) {
            if (data.Length - recordOffset < EmfRecordHeaderSizeBytes) {
                return false;
            }

            int recordType = ReadInt32LittleEndian(data, recordOffset);
            uint recordSize = ReadUInt32LittleEndian(data, recordOffset + 4);
            long recordEnd = (long)recordOffset + recordSize;
            if (recordType <= 0 ||
                recordSize < EmfRecordHeaderSizeBytes ||
                (recordSize & 3) != 0 ||
                recordEnd > data.LongLength) {
                return false;
            }

            recordCount++;
            if (recordType == 14) {
                return recordSize >= EmfEofMinimumSizeBytes &&
                    ReadUInt32LittleEndian(data, (int)recordEnd - 4) == recordSize &&
                    recordEnd == data.LongLength &&
                    recordCount == declaredRecords;
            }

            if (recordCount >= declaredRecords) {
                return false;
            }

            recordOffset = (int)recordEnd;
        }

        return false;
    }
}
