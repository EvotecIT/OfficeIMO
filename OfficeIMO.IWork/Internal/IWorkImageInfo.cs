using System.IO.Compression;
using OfficeIMO.Drawing;

namespace OfficeIMO.IWork.Internal;

internal static class IWorkImageInfo {
    private static readonly uint[] PngCrcTable = CreatePngCrcTable();
    internal static (int? Width, int? Height) Read(byte[] bytes, string mediaType,
        long maximumDecodedBytes) {
        if (string.Equals(mediaType, "image/png", StringComparison.OrdinalIgnoreCase)
            && TryReadPng(bytes, maximumDecodedBytes, out int width, out int height)) {
            return (width, height);
        }
        if (string.Equals(mediaType, "image/jpeg", StringComparison.OrdinalIgnoreCase)) {
            return ReadJpeg(bytes, maximumDecodedBytes);
        }
        return (null, null);
    }

    private static (int? Width, int? Height) ReadJpeg(byte[] bytes, long maximumDecodedBytes) {
        (int? width, int? height) = ReadJpegMetadata(bytes);
        if (!width.HasValue || !height.HasValue
            || (long)width.Value * height.Value * 4 > maximumDecodedBytes
            || bytes.Length < 2 || bytes[bytes.Length - 2] != 0xff || bytes[bytes.Length - 1] != 0xd9
            || !OfficeJpegCodec.TryDecode(bytes, out OfficeRasterImage? decoded)
            || decoded == null || decoded.Width != width.Value || decoded.Height != height.Value) {
            return (null, null);
        }
        return (decoded.Width, decoded.Height);
    }

    private static (int? Width, int? Height) ReadJpegMetadata(byte[] bytes) {
        if (bytes.Length < 4 || bytes[0] != 0xff || bytes[1] != 0xd8) return (null, null);
        int offset = 2;
        int width = 0;
        int height = 0;
        bool hasScan = false;
        while (offset < bytes.Length) {
            if (bytes[offset++] != 0xff) return (null, null);
            if (offset >= bytes.Length) return (null, null);
            byte marker = bytes[offset++];
            while (marker == 0xff && offset < bytes.Length) marker = bytes[offset++];
            if (marker == 0xd9) return hasScan && width > 0 && height > 0
                ? (width, height)
                : (null, null);
            if (marker == 0x00 || marker == 0xd8) return (null, null);
            if (marker == 0x01 || marker >= 0xd0 && marker <= 0xd7) continue;
            if (offset > bytes.Length - 2) return (null, null);
            int length = bytes[offset] << 8 | bytes[offset + 1];
            if (length < 2 || offset > bytes.Length - length) return (null, null);
            bool sizeMarker = marker >= 0xc0 && marker <= 0xc3
                || marker >= 0xc5 && marker <= 0xc7
                || marker >= 0xc9 && marker <= 0xcb
                || marker >= 0xcd && marker <= 0xcf;
            if (sizeMarker && length >= 7) {
                height = bytes[offset + 3] << 8 | bytes[offset + 4];
                width = bytes[offset + 5] << 8 | bytes[offset + 6];
                if (width <= 0 || height <= 0) return (null, null);
            }
            offset += length;
            if (marker == 0xda) {
                hasScan = true;
                if (!SeekNextJpegMarker(bytes, ref offset)) return (null, null);
            }
        }
        return (null, null);
    }

    private static bool SeekNextJpegMarker(byte[] bytes, ref int offset) {
        while (offset < bytes.Length) {
            if (bytes[offset++] != 0xff) continue;
            int markerOffset = offset - 1;
            if (offset >= bytes.Length) return false;
            byte marker = bytes[offset++];
            while (marker == 0xff && offset < bytes.Length) marker = bytes[offset++];
            if (marker == 0x00 || marker >= 0xd0 && marker <= 0xd7) continue;
            offset = markerOffset;
            return true;
        }
        return false;
    }

    private static bool HasPngSignature(byte[] bytes) =>
        bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4e && bytes[3] == 0x47
        && bytes[4] == 0x0d && bytes[5] == 0x0a && bytes[6] == 0x1a && bytes[7] == 0x0a;

    private static bool TryReadPng(byte[] bytes, long maximumDecodedBytes,
        out int width, out int height) {
        width = 0;
        height = 0;
        if (bytes.Length < 33 || !HasPngSignature(bytes)) return false;

        bool hasHeader = false;
        bool hasImageData = false;
        bool imageDataEnded = false;
        bool hasPalette = false;
        int paletteEntryCount = 0;
        byte bitDepth = 0;
        byte colorType = 0;
        byte interlace = 0;
        using var imageData = new MemoryStream();
        int offset = 8;
        while (offset <= bytes.Length - 12) {
            uint rawLength = ReadBigEndianUInt32(bytes, offset);
            if (rawLength > int.MaxValue) return false;
            int dataLength = (int)rawLength;
            if (dataLength > bytes.Length - offset - 12) return false;

            int typeOffset = offset + 4;
            int dataOffset = typeOffset + 4;
            int crcOffset = dataOffset + dataLength;
            uint expectedCrc = ReadBigEndianUInt32(bytes, crcOffset);
            if (expectedCrc != CalculatePngCrc(bytes, typeOffset, checked(4 + dataLength))) return false;

            bool isHeader = IsChunk(bytes, typeOffset, 'I', 'H', 'D', 'R');
            bool isImageData = IsChunk(bytes, typeOffset, 'I', 'D', 'A', 'T');
            bool isEnd = IsChunk(bytes, typeOffset, 'I', 'E', 'N', 'D');
            bool isPalette = IsChunk(bytes, typeOffset, 'P', 'L', 'T', 'E');
            if (!hasHeader) {
                if (!isHeader || offset != 8 || dataLength != 13) return false;
                width = ReadBigEndian32(bytes, dataOffset);
                height = ReadBigEndian32(bytes, dataOffset + 4);
                bitDepth = bytes[dataOffset + 8];
                colorType = bytes[dataOffset + 9];
                interlace = bytes[dataOffset + 12];
                if (width <= 0 || height <= 0
                    || bytes[dataOffset + 10] != 0 || bytes[dataOffset + 11] != 0
                    || interlace > 1 || !IsValidPngColorFormat(bitDepth, colorType)) return false;
                hasHeader = true;
            } else if (isHeader) {
                return false;
            }

            if (isPalette) {
                int maximumPaletteLength = colorType == 3 ? 3 * (1 << bitDepth) : 768;
                if (hasPalette || hasImageData || colorType is 0 or 4
                    || dataLength < 3 || dataLength % 3 != 0
                    || dataLength > maximumPaletteLength) return false;
                hasPalette = true;
                paletteEntryCount = dataLength / 3;
            } else if (isImageData) {
                if (colorType == 3 && !hasPalette) return false;
                if (imageDataEnded || dataLength == 0) return false;
                imageData.Write(bytes, dataOffset, dataLength);
                hasImageData = true;
            } else if (hasImageData && !isEnd) {
                imageDataEnded = true;
            }
            bool isUnknownCritical = (bytes[typeOffset] & 0x20) == 0
                && !isHeader && !isPalette && !isImageData && !isEnd;
            if (isUnknownCritical) return false;
            offset = crcOffset + 4;
            if (isEnd) {
                return dataLength == 0 && hasImageData && offset == bytes.Length
                    && ValidatePngImageData(imageData.ToArray(), width, height,
                        bitDepth, colorType, interlace, paletteEntryCount, maximumDecodedBytes);
            }
        }
        return false;
    }

    private static bool IsValidPngColorFormat(byte bitDepth, byte colorType) => colorType switch {
        0 => bitDepth is 1 or 2 or 4 or 8 or 16,
        2 => bitDepth is 8 or 16,
        3 => bitDepth is 1 or 2 or 4 or 8,
        4 => bitDepth is 8 or 16,
        6 => bitDepth is 8 or 16,
        _ => false
    };

    private static bool ValidatePngImageData(byte[] data, int width, int height,
        byte bitDepth, byte colorType, byte interlace, int paletteEntryCount,
        long maximumDecodedBytes) {
        if (data.Length < 6) return false;
        byte compressionMethod = (byte)(data[0] & 0x0f);
        int windowSize = data[0] >> 4;
        bool hasPresetDictionary = (data[1] & 0x20) != 0;
        if (compressionMethod != 8 || windowSize > 7 || hasPresetDictionary
            || ((data[0] << 8) | data[1]) % 31 != 0) return false;

        long decodedLength;
        try {
            decodedLength = ExpectedPngDataLength(width, height, bitDepth, colorType, interlace);
        } catch (OverflowException) {
            return false;
        }
        if (decodedLength <= 0 || decodedLength > maximumDecodedBytes) return false;

        uint expectedAdler = ReadBigEndianUInt32(data, data.Length - 4);
        using var compressed = new MemoryStream(data, 2, data.Length - 6, writable: false);
        try {
            using var inflater = new DeflateStream(compressed, CompressionMode.Decompress, leaveOpen: true);
            uint first = 1;
            uint second = 0;
            long decoded = 0;
            var buffer = new byte[8192];
            foreach ((int passWidth, int passHeight) in PngPasses(width, height, interlace)) {
                int channels = colorType switch { 0 or 3 => 1, 2 => 3, 4 => 2, 6 => 4, _ => 0 };
                long rowBytes = checked(((long)passWidth * channels * bitDepth + 7) / 8);
                byte[]? previousIndexedRow = null;
                byte[]? indexedRow = null;
                if (colorType == 3) {
                    if (rowBytes > int.MaxValue
                        || decodedLength > maximumDecodedBytes - checked(rowBytes * 2)) return false;
                    previousIndexedRow = new byte[(int)rowBytes];
                    indexedRow = new byte[(int)rowBytes];
                }
                for (int row = 0; row < passHeight; row++) {
                    int filter = ReadPngByte(inflater, ref first, ref second, ref decoded);
                    if (filter is < 0 or > 4) return false;
                    if (indexedRow != null && previousIndexedRow != null) {
                        if (!ReadPngBytes(inflater, indexedRow, indexedRow.Length,
                                ref first, ref second, ref decoded)
                            || !UnfilterIndexedPngRow(indexedRow, previousIndexedRow, filter)
                            || !IndexedPngRowUsesPalette(indexedRow, passWidth,
                                bitDepth, paletteEntryCount)) return false;
                        (previousIndexedRow, indexedRow) = (indexedRow, previousIndexedRow);
                    } else if (!ReadPngBytes(inflater, buffer, rowBytes,
                                   ref first, ref second, ref decoded)) return false;
                }
            }
            if (decoded != decodedLength || inflater.ReadByte() != -1) return false;
            return ((second << 16) | first) == expectedAdler;
        } catch (Exception exception) when (exception is InvalidDataException or IOException) {
            return false;
        }
    }

    private static bool UnfilterIndexedPngRow(byte[] row, byte[] previous, int filter) {
        for (int index = 0; index < row.Length; index++) {
            byte left = index > 0 ? row[index - 1] : (byte)0;
            byte above = previous[index];
            byte upperLeft = index > 0 ? previous[index - 1] : (byte)0;
            int reconstructed = filter switch {
                0 => row[index],
                1 => row[index] + left,
                2 => row[index] + above,
                3 => row[index] + ((left + above) >> 1),
                4 => row[index] + PaethPredictor(left, above, upperLeft),
                _ => -1
            };
            if (reconstructed < 0) return false;
            row[index] = unchecked((byte)reconstructed);
        }
        return true;
    }

    private static int PaethPredictor(int left, int above, int upperLeft) {
        int estimate = left + above - upperLeft;
        int leftDistance = Math.Abs(estimate - left);
        int aboveDistance = Math.Abs(estimate - above);
        int upperLeftDistance = Math.Abs(estimate - upperLeft);
        return leftDistance <= aboveDistance && leftDistance <= upperLeftDistance
            ? left
            : aboveDistance <= upperLeftDistance ? above : upperLeft;
    }

    private static bool IndexedPngRowUsesPalette(byte[] row, int width,
        byte bitDepth, int paletteEntryCount) {
        int mask = (1 << bitDepth) - 1;
        for (int pixel = 0; pixel < width; pixel++) {
            int bitOffset = pixel * bitDepth;
            int shift = 8 - bitDepth - bitOffset % 8;
            int index = row[bitOffset / 8] >> shift & mask;
            if (index >= paletteEntryCount) return false;
        }
        return true;
    }

    private static long ExpectedPngDataLength(int width, int height, byte bitDepth,
        byte colorType, byte interlace) {
        int channels = colorType switch { 0 or 3 => 1, 2 => 3, 4 => 2, 6 => 4, _ => 0 };
        long total = 0;
        foreach ((int passWidth, int passHeight) in PngPasses(width, height, interlace)) {
            long rowBytes = checked(((long)passWidth * channels * bitDepth + 7) / 8);
            total = checked(total + checked((rowBytes + 1) * passHeight));
        }
        return total;
    }

    private static IEnumerable<(int Width, int Height)> PngPasses(int width, int height,
        byte interlace) {
        if (interlace == 0) {
            yield return (width, height);
            yield break;
        }
        int[] xStart = { 0, 4, 0, 2, 0, 1, 0 };
        int[] yStart = { 0, 0, 4, 0, 2, 0, 1 };
        int[] xStep = { 8, 8, 4, 4, 2, 2, 1 };
        int[] yStep = { 8, 8, 8, 4, 4, 2, 2 };
        for (int pass = 0; pass < xStart.Length; pass++) {
            int passWidth = width <= xStart[pass]
                ? 0 : checked((int)(((long)width - xStart[pass] + xStep[pass] - 1) / xStep[pass]));
            int passHeight = height <= yStart[pass]
                ? 0 : checked((int)(((long)height - yStart[pass] + yStep[pass] - 1) / yStep[pass]));
            if (passWidth > 0 && passHeight > 0) yield return (passWidth, passHeight);
        }
    }

    private static int ReadPngByte(Stream stream, ref uint first, ref uint second,
        ref long decoded) {
        int value = stream.ReadByte();
        if (value < 0) return -1;
        UpdateAdler((byte)value, ref first, ref second);
        decoded++;
        return value;
    }

    private static bool ReadPngBytes(Stream stream, byte[] buffer, long count,
        ref uint first, ref uint second, ref long decoded) {
        while (count > 0) {
            int requested = (int)Math.Min(buffer.Length, count);
            int read = stream.Read(buffer, 0, requested);
            if (read <= 0) return false;
            for (int index = 0; index < read; index++) {
                UpdateAdler(buffer[index], ref first, ref second);
            }
            decoded += read;
            count -= read;
        }
        return true;
    }

    private static void UpdateAdler(byte value, ref uint first, ref uint second) {
        const uint modulus = 65521;
        first += value;
        if (first >= modulus) first -= modulus;
        second += first;
        if (second >= modulus) second -= modulus;
    }

    private static bool IsChunk(byte[] bytes, int offset, char first, char second, char third, char fourth) =>
        bytes[offset] == first && bytes[offset + 1] == second
        && bytes[offset + 2] == third && bytes[offset + 3] == fourth;

    private static uint CalculatePngCrc(byte[] bytes, int offset, int length) {
        uint crc = uint.MaxValue;
        for (int index = offset; index < offset + length; index++) {
            crc = PngCrcTable[(int)((crc ^ bytes[index]) & 0xff)] ^ crc >> 8;
        }
        return crc ^ uint.MaxValue;
    }

    private static uint[] CreatePngCrcTable() {
        var table = new uint[256];
        for (uint value = 0; value < table.Length; value++) {
            uint crc = value;
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1) != 0 ? 0xedb88320U ^ crc >> 1 : crc >> 1;
            }
            table[(int)value] = crc;
        }
        return table;
    }

    private static int ReadBigEndian32(byte[] bytes, int offset) =>
        bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3];

    private static uint ReadBigEndianUInt32(byte[] bytes, int offset) =>
        (uint)bytes[offset] << 24 | (uint)bytes[offset + 1] << 16
        | (uint)bytes[offset + 2] << 8 | bytes[offset + 3];
}
