namespace OfficeIMO.Drawing;

public static partial class OfficeImageReader {
    private const int ClassicTiffMagic = 42;
    private const int BigTiffMagic = 43;

    private static bool TryReadTiff(byte[] data, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 8) {
            return false;
        }

        bool littleEndian;
        if (data[0] == (byte)'I' && data[1] == (byte)'I') {
            littleEndian = true;
        } else if (data[0] == (byte)'M' && data[1] == (byte)'M') {
            littleEndian = false;
        } else {
            return false;
        }

        return ReadUInt16(data, 2, littleEndian) switch {
            ClassicTiffMagic => TryReadClassicTiff(data, littleEndian, out info),
            BigTiffMagic => TryReadBigTiff(data, littleEndian, out info),
            _ => false
        };
    }

    private static bool TryReadClassicTiff(byte[] data, bool littleEndian, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        int ifdOffset = ReadInt32(data, 4, littleEndian);
        if (ifdOffset < 0 || ifdOffset > data.Length - 2) {
            return false;
        }

        int entryCount = ReadUInt16(data, ifdOffset, littleEndian);
        int width = 0;
        int height = 0;
        double dpiX = 96.0;
        double dpiY = 96.0;
        int unit = 2;

        for (int i = 0; i < entryCount; i++) {
            int entry = ifdOffset + 2 + (i * 12);
            if (entry < 0 || entry > data.Length - 12) {
                break;
            }

            int tag = ReadUInt16(data, entry, littleEndian);
            int type = ReadUInt16(data, entry + 2, littleEndian);
            int count = ReadInt32(data, entry + 4, littleEndian);
            int valueOrOffset = ReadInt32(data, entry + 8, littleEndian);

            if (tag == 256) width = ReadClassicTiffScalar(type, count, valueOrOffset, littleEndian);
            else if (tag == 257) height = ReadClassicTiffScalar(type, count, valueOrOffset, littleEndian);
            else if (tag == 282) dpiX = ReadClassicTiffRational(data, valueOrOffset, littleEndian, dpiX);
            else if (tag == 283) dpiY = ReadClassicTiffRational(data, valueOrOffset, littleEndian, dpiY);
            else if (tag == 296) unit = ReadClassicTiffScalar(type, count, valueOrOffset, littleEndian);
        }

        return CompleteTiffInfo(width, height, dpiX, dpiY, unit, out info);
    }

    private static bool TryReadBigTiff(byte[] data, bool littleEndian, out OfficeImageInfo info) {
        info = new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
        if (data.Length < 16 ||
            ReadUInt16(data, 4, littleEndian) != 8 ||
            ReadUInt16(data, 6, littleEndian) != 0) {
            return false;
        }

        ulong ifdOffsetValue = ReadUInt64(data, 8, littleEndian);
        if (ifdOffsetValue > int.MaxValue) {
            return false;
        }

        int ifdOffset = (int)ifdOffsetValue;
        if (ifdOffset < 16 || ifdOffset > data.Length - 8) {
            return false;
        }

        ulong declaredEntryCount = ReadUInt64(data, ifdOffset, littleEndian);
        int availableEntryCount = (data.Length - ifdOffset - 8) / 20;
        int entryCount = declaredEntryCount > (ulong)availableEntryCount
            ? availableEntryCount
            : (int)declaredEntryCount;
        int width = 0;
        int height = 0;
        double dpiX = 96.0;
        double dpiY = 96.0;
        int unit = 2;

        for (int i = 0; i < entryCount; i++) {
            int entry = ifdOffset + 8 + (i * 20);
            int tag = ReadUInt16(data, entry, littleEndian);
            int type = ReadUInt16(data, entry + 2, littleEndian);
            ulong count = ReadUInt64(data, entry + 4, littleEndian);

            if (tag == 256) width = ReadBigTiffScalar(data, entry + 12, type, count, littleEndian);
            else if (tag == 257) height = ReadBigTiffScalar(data, entry + 12, type, count, littleEndian);
            else if (tag == 282) dpiX = ReadBigTiffRational(data, entry + 12, type, count, littleEndian, dpiX);
            else if (tag == 283) dpiY = ReadBigTiffRational(data, entry + 12, type, count, littleEndian, dpiY);
            else if (tag == 296) unit = ReadBigTiffScalar(data, entry + 12, type, count, littleEndian);
        }

        return CompleteTiffInfo(width, height, dpiX, dpiY, unit, out info);
    }

    private static bool CompleteTiffInfo(
        int width,
        int height,
        double dpiX,
        double dpiY,
        int unit,
        out OfficeImageInfo info) {
        switch (unit) {
            case 2:
                break;
            case 3:
                dpiX *= 2.54;
                dpiY *= 2.54;
                break;
            default:
                // ResolutionUnit=None (1) describes unitless ratios, not DPI.
                // Unknown values likewise cannot be exported as physical resolution.
                dpiX = 96.0;
                dpiY = 96.0;
                break;
        }

        info = new OfficeImageInfo(OfficeImageFormat.Tiff, width, height, dpiX, dpiY);
        return width > 0 && height > 0;
    }

    private static int ReadClassicTiffScalar(int type, int count, int valueOrOffset, bool littleEndian) {
        if (count <= 0) return 0;
        if (type == 3) {
            return littleEndian ? valueOrOffset & 0xFFFF : (valueOrOffset >> 16) & 0xFFFF;
        }

        return type == 4 ? valueOrOffset : 0;
    }

    private static double ReadClassicTiffRational(byte[] data, int offset, bool littleEndian, double fallback) {
        if (offset < 0 || offset > data.Length - 8) return fallback;
        uint numerator = ReadUInt32(data, offset, littleEndian);
        uint denominator = ReadUInt32(data, offset + 4, littleEndian);
        return denominator != 0 ? (double)numerator / denominator : fallback;
    }

    private static int ReadBigTiffScalar(byte[] data, int offset, int type, ulong count, bool littleEndian) {
        if (count != 1) return 0;
        ulong value = type switch {
            3 => (ulong)ReadUInt16(data, offset, littleEndian),
            4 => ReadUInt32(data, offset, littleEndian),
            16 => ReadUInt64(data, offset, littleEndian),
            _ => 0
        };
        return value <= int.MaxValue ? (int)value : 0;
    }

    private static double ReadBigTiffRational(
        byte[] data,
        int offset,
        int type,
        ulong count,
        bool littleEndian,
        double fallback) {
        if (type != 5 || count != 1) return fallback;
        uint numerator = ReadUInt32(data, offset, littleEndian);
        uint denominator = ReadUInt32(data, offset + 4, littleEndian);
        return denominator != 0 ? (double)numerator / denominator : fallback;
    }

    private static uint ReadUInt32(byte[] data, int offset, bool littleEndian) {
        if (offset < 0 || offset > data.Length - 4) return 0;
        return littleEndian
            ? (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24))
            : (uint)((data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]);
    }

    private static ulong ReadUInt64(byte[] data, int offset, bool littleEndian) {
        if (offset < 0 || offset > data.Length - 8) return 0;
        ulong value = 0;
        if (littleEndian) {
            for (int index = 7; index >= 0; index--) {
                value = (value << 8) | data[offset + index];
            }
        } else {
            for (int index = 0; index < 8; index++) {
                value = (value << 8) | data[offset + index];
            }
        }
        return value;
    }
}
