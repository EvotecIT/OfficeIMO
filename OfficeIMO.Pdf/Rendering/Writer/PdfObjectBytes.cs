using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfObjectBytes {
    internal static byte[] WrapIndirectObject(int objectNumber, string body) {
        Guard.NotNull(body, nameof(body));
        return WrapIndirectObject(objectNumber, PdfEncoding.Latin1GetBytes(body));
    }

    internal static byte[] WrapIndirectObject(int objectNumber, byte[] body) {
        return WrapIndirectObject(objectNumber, 0, body);
    }

    internal static byte[] WrapIndirectObject(int objectNumber, int generation, byte[] body) {
        Guard.NotNull(body, nameof(body));
        if (objectNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        }

        if (generation < 0) {
            throw new ArgumentOutOfRangeException(nameof(generation), "PDF object generation cannot be negative.");
        }

        return Concat(
            PdfEncoding.Latin1GetBytes(objectNumber.ToString(CultureInfo.InvariantCulture) + " " + generation.ToString(CultureInfo.InvariantCulture) + " obj\n"),
            body,
            PdfEncoding.Latin1GetBytes("endobj\n"));
    }

    internal static byte[] WrapStreamObject(int objectNumber, string dictionary, byte[] content) {
        return WrapIndirectObject(objectNumber, WrapStreamBody(dictionary, content));
    }

    internal static byte[] WrapStreamBody(string dictionary, byte[] content) {
        Guard.NotNull(content, nameof(content));
        Guard.NotNullOrWhiteSpace(dictionary, nameof(dictionary));
        if (ContainsStreamMarker(dictionary)) {
            throw new ArgumentException("Stream dictionaries must not include stream markers.", nameof(dictionary));
        }

        return Concat(
            PdfEncoding.Latin1GetBytes(dictionary.TrimEnd() + "\nstream\n"),
            content,
            PdfEncoding.Latin1GetBytes("\nendstream\n"));
    }

    private static bool ContainsStreamMarker(string dictionary) {
        int index = 0;
        while ((index = dictionary.IndexOf("stream", index, StringComparison.Ordinal)) >= 0) {
            bool before = index == 0 || char.IsWhiteSpace(dictionary[index - 1]);
            int afterIndex = index + 6;
            bool after = afterIndex == dictionary.Length || char.IsWhiteSpace(dictionary[afterIndex]);
            if (before && after) return true;
            index = afterIndex;
        }
        return false;
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
