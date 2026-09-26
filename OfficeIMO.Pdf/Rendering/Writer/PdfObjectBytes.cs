using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static class PdfObjectBytes {
    internal static byte[] WrapIndirectObject(int objectNumber, string body) {
        Guard.NotNull(body, nameof(body));
        return WrapIndirectObject(objectNumber, PdfEncoding.Latin1GetBytes(body));
    }

    internal static byte[] WrapIndirectObject(int objectNumber, byte[] body) {
        return WrapIndirectObject(objectNumber, 0, body);
    }

    internal static byte[] WrapIndirectObjectCancellable(int objectNumber, byte[] body, CancellationToken cancellationToken) {
        Guard.NotNull(body, nameof(body));
        if (objectNumber < 1) throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        cancellationToken.ThrowIfCancellationRequested();
        byte[] prefix = PdfEncoding.Latin1GetBytes(objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
        byte[] suffix = PdfEncoding.Latin1GetBytes("endobj\n");
        var result = new byte[checked(prefix.Length + body.Length + suffix.Length)];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        for (int offset = 0; offset < body.Length; offset += 65536) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(body, offset, result, prefix.Length + offset, Math.Min(65536, body.Length - offset));
        }
        cancellationToken.ThrowIfCancellationRequested();
        Buffer.BlockCopy(suffix, 0, result, prefix.Length + body.Length, suffix.Length);
        return result;
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
        return Concat(CreateStreamObjectSegments(objectNumber, dictionary, content));
    }

    internal static byte[] WrapStreamObject(int objectNumber, string dictionary, PdfStream content, CancellationToken cancellationToken = default) {
        Guard.NotNull(content, nameof(content));
        CreateStreamObjectEnvelope(objectNumber, dictionary, out byte[] prefix, out byte[] suffix, cancellationToken);
        var result = new byte[checked(prefix.Length + content.DataLength + suffix.Length)];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        content.CopyDataTo(result, prefix.Length, cancellationToken);
        Buffer.BlockCopy(suffix, 0, result, prefix.Length + content.DataLength, suffix.Length);
        return result;
    }

    internal static PdfSerializedObject SegmentStreamObject(int objectNumber, string dictionary, PdfStream content, CancellationToken cancellationToken = default) {
        Guard.NotNull(content, nameof(content));
        CreateStreamObjectEnvelope(objectNumber, dictionary, out byte[] prefix, out byte[] suffix, cancellationToken);
        return PdfSerializedObject.FromStream(prefix, content, suffix);
    }

    internal static byte[][] CreateStreamObjectSegments(int objectNumber, string dictionary, byte[] content) {
        Guard.NotNull(content, nameof(content));
        CreateStreamObjectEnvelope(objectNumber, dictionary, out byte[] prefix, out byte[] suffix);
        return new[] {
            prefix,
            content,
            suffix
        };
    }

    private static void CreateStreamObjectEnvelope(
        int objectNumber,
        string dictionary,
        out byte[] prefix,
        out byte[] suffix,
        CancellationToken cancellationToken = default) {
        ValidateStreamDictionary(dictionary, cancellationToken);
        if (objectNumber < 1) throw new ArgumentOutOfRangeException(nameof(objectNumber), "PDF object number must be positive.");
        if (ContainsStreamMarker(dictionary, cancellationToken)) throw new ArgumentException("Stream dictionaries must not include stream markers.", nameof(dictionary));

        string header = objectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n";
        prefix = CreateStreamPrefix(header, dictionary, cancellationToken);
        suffix = PdfEncoding.Latin1GetBytes("\nendstream\nendobj\n");
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

    internal static byte[] WrapStreamBody(string dictionary, PdfStream content, CancellationToken cancellationToken = default) {
        Guard.NotNull(content, nameof(content));
        ValidateStreamDictionary(dictionary, cancellationToken);
        if (ContainsStreamMarker(dictionary, cancellationToken)) {
            throw new ArgumentException("Stream dictionaries must not include stream markers.", nameof(dictionary));
        }

        byte[] prefix = CreateStreamPrefix(string.Empty, dictionary, cancellationToken);
        byte[] suffix = PdfEncoding.Latin1GetBytes("\nendstream\n");
        var result = new byte[checked(prefix.Length + content.DataLength + suffix.Length)];
        Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
        content.CopyDataTo(result, prefix.Length, cancellationToken);
        Buffer.BlockCopy(suffix, 0, result, prefix.Length + content.DataLength, suffix.Length);
        return result;
    }

    private static byte[] CreateStreamPrefix(string header, string dictionary, CancellationToken cancellationToken) {
        int dictionaryLength = dictionary.Length;
        while (dictionaryLength > 0 && char.IsWhiteSpace(dictionary[dictionaryLength - 1])) {
            cancellationToken.ThrowIfCancellationRequested();
            dictionaryLength--;
        }
        const string streamStart = "\nstream\n";
        byte[] prefix = new byte[checked(header.Length + dictionaryLength + streamStart.Length)];
        int offset = 0;
        for (int i = 0; i < header.Length; i++) prefix[offset++] = (byte)header[i];
        for (int i = 0; i < dictionaryLength; i++) {
            if ((i & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            prefix[offset++] = (byte)(dictionary[i] & 0xFF);
        }
        cancellationToken.ThrowIfCancellationRequested();
        for (int i = 0; i < streamStart.Length; i++) prefix[offset++] = (byte)streamStart[i];
        return prefix;
    }

    private static void ValidateStreamDictionary(string dictionary, CancellationToken cancellationToken) {
        Guard.NotNull(dictionary, nameof(dictionary));
        bool hasContent = false;
        for (int index = 0; index < dictionary.Length; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (!char.IsWhiteSpace(dictionary[index])) hasContent = true;
        }
        if (!hasContent) throw new ArgumentException("Value cannot be null or whitespace.", nameof(dictionary));
    }

    private static bool ContainsStreamMarker(string dictionary, CancellationToken cancellationToken = default) {
        for (int index = 0; index <= dictionary.Length - 6; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (dictionary[index] != 's' || dictionary[index + 1] != 't' ||
                dictionary[index + 2] != 'r' || dictionary[index + 3] != 'e' ||
                dictionary[index + 4] != 'a' || dictionary[index + 5] != 'm') continue;
            bool before = index == 0 || char.IsWhiteSpace(dictionary[index - 1]);
            int afterIndex = index + 6;
            bool after = afterIndex == dictionary.Length || char.IsWhiteSpace(dictionary[afterIndex]);
            if (before && after) return true;
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
