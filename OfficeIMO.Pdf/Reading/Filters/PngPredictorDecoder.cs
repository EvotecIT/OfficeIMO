namespace OfficeIMO.Pdf.Filters;

internal static class PngPredictorDecoder {
    public static byte[] Decode(byte[] data, int columns, int colors, int bitsPerComponent, int maxOutputBytes) {
        if (maxOutputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxOutputBytes), maxOutputBytes, "Maximum decoded stream bytes must be positive.");
        }

        if (data == null || data.Length == 0) {
            return Array.Empty<byte>();
        }

        colors = Math.Max(1, colors);
        bitsPerComponent = Math.Max(1, bitsPerComponent);
        columns = Math.Max(1, columns);

        long bitsPerPixel = ((long)colors * bitsPerComponent) + 7L;
        long rowBits;
        try {
            rowBits = checked((long)columns * colors * bitsPerComponent);
        } catch (OverflowException) {
            throw CreateDecodedLimitException(maxOutputBytes, (long)maxOutputBytes + 1L);
        }

        long rowLengthValue = (rowBits + 7L) / 8L;
        if (rowLengthValue <= 0L || rowLengthValue > maxOutputBytes) {
            throw CreateDecodedLimitException(maxOutputBytes, Math.Max(rowLengthValue, (long)maxOutputBytes + 1L));
        }

        int rowLength = (int)rowLengthValue;
        int bytesPerPixel = (int)Math.Max(1L, bitsPerPixel / 8L);
        long encodedRowLength = rowLengthValue + 1L;
        if (data.LongLength % encodedRowLength != 0L) {
            throw new FormatException("PNG predictor row exceeds decoded stream length.");
        }

        long outputLength = (data.LongLength / encodedRowLength) * rowLengthValue;
        if (outputLength > maxOutputBytes) {
            throw CreateDecodedLimitException(maxOutputBytes, outputLength);
        }

        var output = new byte[(int)outputLength];
        int outputOffset = 0;
        int inputOffset = 0;
        var previousRow = new byte[rowLength];
        var currentRow = new byte[rowLength];

        while (inputOffset < data.Length) {
            int filterType = data[inputOffset++];
            if (inputOffset + rowLength > data.Length) {
                throw new FormatException("PNG predictor row exceeds decoded stream length.");
            }

            Buffer.BlockCopy(data, inputOffset, currentRow, 0, rowLength);
            inputOffset += rowLength;

            switch (filterType) {
                case 0:
                    break;
                case 1:
                    for (int i = 0; i < rowLength; i++) {
                        int left = i >= bytesPerPixel ? currentRow[i - bytesPerPixel] : 0;
                        currentRow[i] = unchecked((byte)(currentRow[i] + left));
                    }
                    break;
                case 2:
                    for (int i = 0; i < rowLength; i++) {
                        currentRow[i] = unchecked((byte)(currentRow[i] + previousRow[i]));
                    }
                    break;
                case 3:
                    for (int i = 0; i < rowLength; i++) {
                        int left = i >= bytesPerPixel ? currentRow[i - bytesPerPixel] : 0;
                        int up = previousRow[i];
                        currentRow[i] = unchecked((byte)(currentRow[i] + ((left + up) / 2)));
                    }
                    break;
                case 4:
                    for (int i = 0; i < rowLength; i++) {
                        int left = i >= bytesPerPixel ? currentRow[i - bytesPerPixel] : 0;
                        int up = previousRow[i];
                        int upLeft = i >= bytesPerPixel ? previousRow[i - bytesPerPixel] : 0;
                        currentRow[i] = unchecked((byte)(currentRow[i] + PaethPredictor(left, up, upLeft)));
                    }
                    break;
                default:
                    throw new FormatException($"Unsupported PNG predictor filter type '{filterType}'.");
            }

            Buffer.BlockCopy(currentRow, 0, output, outputOffset, rowLength);
            outputOffset += rowLength;
            Buffer.BlockCopy(currentRow, 0, previousRow, 0, rowLength);
        }

        if (outputOffset == output.Length) {
            return output;
        }

        var trimmed = new byte[outputOffset];
        Buffer.BlockCopy(output, 0, trimmed, 0, outputOffset);
        return trimmed;
    }

    private static PdfReadLimitException CreateDecodedLimitException(int maximum, long actual) =>
        PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximum, actual);

    private static int PaethPredictor(int left, int up, int upLeft) {
        int prediction = left + up - upLeft;
        int distanceLeft = Math.Abs(prediction - left);
        int distanceUp = Math.Abs(prediction - up);
        int distanceUpLeft = Math.Abs(prediction - upLeft);

        if (distanceLeft <= distanceUp && distanceLeft <= distanceUpLeft) {
            return left;
        }

        return distanceUp <= distanceUpLeft ? up : upLeft;
    }
}
