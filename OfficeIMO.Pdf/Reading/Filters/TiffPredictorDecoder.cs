namespace OfficeIMO.Pdf.Filters;

internal static class TiffPredictorDecoder {
    public static byte[] Decode(byte[] data, int columns, int colors = 1, int bitsPerComponent = 8) {
        if (data == null || data.Length == 0) {
            return Array.Empty<byte>();
        }

        colors = Math.Max(1, colors);
        columns = Math.Max(1, columns);
        bitsPerComponent = Math.Max(1, bitsPerComponent);

        if (bitsPerComponent != 8) {
            return data;
        }

        int rowLength = columns * colors;
        if (rowLength <= 0 || data.Length % rowLength != 0) {
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
}
