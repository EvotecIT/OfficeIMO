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

        colors = Math.Max(1, colors);
        columns = Math.Max(1, columns);
        bitsPerComponent = Math.Max(1, bitsPerComponent);

        if (bitsPerComponent != 8) {
            return data;
        }

        long rowLengthValue = (long)columns * colors;
        if (rowLengthValue <= 0L || rowLengthValue > maxOutputBytes) {
            throw CreateDecodedLimitException(maxOutputBytes, Math.Max(rowLengthValue, (long)maxOutputBytes + 1L));
        }

        int rowLength = (int)rowLengthValue;
        if (data.Length % rowLength != 0) {
            return data;
        }

        var output = new byte[data.Length];
        for (int rowOffset = 0; rowOffset < data.Length; rowOffset += rowLength) {
            for (int i = 0; i < rowLength; i++) {
                int left = i >= colors ? output[rowOffset + i - colors] : 0;
                output[rowOffset + i] = unchecked((byte)(data[rowOffset + i] + left));
            }
        }

        return output;
    }

    private static PdfReadLimitException CreateDecodedLimitException(int maximum, long actual) =>
        PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximum, actual);
}
