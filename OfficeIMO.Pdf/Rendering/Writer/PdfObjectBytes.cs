using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfObjectBytes {
    internal static byte[] WrapIndirectObject(int objectNumber, string body) {
        Guard.NotNull(body, nameof(body));
        return WrapIndirectObject(objectNumber, PdfEncoding.Latin1GetBytes(body));
    }

    internal static byte[] WrapIndirectObject(int objectNumber, byte[] body) {
        Guard.NotNull(body, nameof(body));
        if (objectNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        }

        return Concat(
            PdfEncoding.Latin1GetBytes(objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n"),
            body,
            PdfEncoding.Latin1GetBytes("endobj\n"));
    }

    internal static byte[] WrapStreamObject(int objectNumber, string dictionary, byte[] content) {
        return WrapIndirectObject(objectNumber, WrapStreamBody(dictionary, content));
    }

    internal static byte[] WrapStreamBody(string dictionary, byte[] content) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNullOrWhiteSpace(dictionary, nameof(dictionary));
        if (dictionary.Contains("stream")) {
            throw new ArgumentException("Stream dictionaries must not include stream markers.", nameof(dictionary));
        }

        return Concat(
            PdfEncoding.Latin1GetBytes(dictionary.TrimEnd() + "\nstream\n"),
            content,
            PdfEncoding.Latin1GetBytes("\nendstream\n"));
    }

    internal static byte[] Concat(params byte[][] parts) {
        int length = 0;
        foreach (byte[] part in parts) {
            length += part.Length;
        }

        var result = new byte[length];
        int offset = 0;
        foreach (byte[] part in parts) {
            Buffer.BlockCopy(part, 0, result, offset, part.Length);
            offset += part.Length;
        }

        return result;
    }
}
