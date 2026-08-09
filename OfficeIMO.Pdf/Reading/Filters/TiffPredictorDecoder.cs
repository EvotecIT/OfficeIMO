namespace OfficeIMO.Pdf.Filters;

internal static class TiffPredictorDecoder {
    public static byte[] Decode(byte[] data, int columns, int colors, int bitsPerComponent, int maxOutputBytes) {
        if (maxOutputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxOutputBytes), maxOutputBytes, "Maximum decoded stream bytes must be positive.");
        }

        if (data == null || data.Length == 0) {
            return Array.Empty<byte>();
        }
        if (data.LongLength > maxOutputBytes) {
            throw CreateDecodedLimitException(maxOutputBytes, data.LongLength);
        }

        if (colors <= 0 || columns <= 0) {
            throw new FormatException("TIFF predictor colors and columns must be positive.");
        }
        if (bitsPerComponent != 1 && bitsPerComponent != 2 && bitsPerComponent != 4 && bitsPerComponent != 8 && bitsPerComponent != 16) {
            throw new FormatException($"Unsupported TIFF predictor bit depth '{bitsPerComponent}'.");
        }

        long samplesPerRow;
        long rowBits;
        try {
            samplesPerRow = checked((long)columns * colors);
            rowBits = checked(samplesPerRow * bitsPerComponent);
        } catch (OverflowException) {
            throw CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
        }

        long rowLengthValue = (rowBits + 7L) / 8L;
        if (rowLengthValue <= 0L || rowLengthValue > maxOutputBytes) {
            throw CreateDecodedLimitException(maxOutputBytes, Math.Max(rowLengthValue, (long)maxOutputBytes + 1L));
        }
        if (samplesPerRow > int.MaxValue) {
            throw CreateDecodedLimitException(maxOutputBytes, Math.Max(rowLengthValue, (long)maxOutputBytes + 1L));
        }

        int rowLength = (int)rowLengthValue;
        if (data.Length % rowLength != 0) {
            throw new FormatException("TIFF predictor data contains an incomplete row.");
        }

        var output = (byte[])data.Clone();
        int sampleCount = checked((int)samplesPerRow);
        int sampleMask = bitsPerComponent == 16 ? ushort.MaxValue : (1 << bitsPerComponent) - 1;
        for (int rowOffset = 0; rowOffset < data.Length; rowOffset += rowLength) {
            if (bitsPerComponent == 8) {
                DecodeEightBitRow(data, output, rowOffset, rowLength, colors);
                continue;
            }
            if (bitsPerComponent == 16) {
                DecodeSixteenBitRow(data, output, rowOffset, sampleCount, colors);
                continue;
            }

            for (int sampleIndex = 0; sampleIndex < sampleCount; sampleIndex++) {
                int encoded = ReadSample(data, rowOffset, sampleIndex, bitsPerComponent);
                int left = sampleIndex >= colors
                    ? ReadSample(output, rowOffset, sampleIndex - colors, bitsPerComponent)
                    : 0;
                WriteSample(output, rowOffset, sampleIndex, bitsPerComponent, (encoded + left) & sampleMask);
            }
        }

        return output;
    }

    private static void DecodeEightBitRow(byte[] input, byte[] output, int rowOffset, int rowLength, int colors) {
        for (int index = 0; index < rowLength; index++) {
            int left = index >= colors ? output[rowOffset + index - colors] : 0;
            output[rowOffset + index] = unchecked((byte)(input[rowOffset + index] + left));
        }
    }

    private static void DecodeSixteenBitRow(byte[] input, byte[] output, int rowOffset, int sampleCount, int colors) {
        for (int sampleIndex = 0; sampleIndex < sampleCount; sampleIndex++) {
            int byteOffset = rowOffset + (sampleIndex * 2);
            int encoded = (input[byteOffset] << 8) | input[byteOffset + 1];
            int left = 0;
            if (sampleIndex >= colors) {
                int leftOffset = rowOffset + ((sampleIndex - colors) * 2);
                left = (output[leftOffset] << 8) | output[leftOffset + 1];
            }

            int decoded = (encoded + left) & ushort.MaxValue;
            output[byteOffset] = (byte)(decoded >> 8);
            output[byteOffset + 1] = (byte)decoded;
        }
    }

    private static int ReadSample(byte[] data, int rowOffset, int sampleIndex, int bitsPerComponent) {
        long bitOffset = (long)sampleIndex * bitsPerComponent;
        int value = 0;
        for (int bit = 0; bit < bitsPerComponent; bit++) {
            long absoluteBit = bitOffset + bit;
            int source = data[rowOffset + (int)(absoluteBit / 8L)];
            value = (value << 1) | ((source >> (7 - (int)(absoluteBit % 8L))) & 1);
        }

        return value;
    }

    private static void WriteSample(byte[] data, int rowOffset, int sampleIndex, int bitsPerComponent, int value) {
        long bitOffset = (long)sampleIndex * bitsPerComponent;
        for (int bit = 0; bit < bitsPerComponent; bit++) {
            long absoluteBit = bitOffset + bit;
            int byteIndex = rowOffset + (int)(absoluteBit / 8L);
            int mask = 1 << (7 - (int)(absoluteBit % 8L));
            int sourceMask = 1 << (bitsPerComponent - bit - 1);
            data[byteIndex] = (byte)((data[byteIndex] & ~mask) | ((value & sourceMask) != 0 ? mask : 0));
        }
    }

    private static PdfReadLimitException CreateDecodedLimitException(int maximum, long actual) =>
        PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximum, actual);
}
